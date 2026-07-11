from __future__ import annotations

import py_compile
import re
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
APP = ROOT / "audit_app.py"
CHANGELOG = ROOT / "CHANGELOG_CUSTOMER.md"
PORTFOLIO = ROOT / "vendor_matrix_detailed.xlsx"


def assert_true(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def check_compile() -> None:
    py_compile.compile(str(APP), doraise=True)


def check_version() -> None:
    text = read_text(APP)
    match = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', text)
    assert_true(match is not None, "APP_VERSION is missing")
    assert_true(match.group(1) == "12.0-dev", f"Unexpected APP_VERSION: {match.group(1)}")


def check_customer_changelog() -> None:
    text = read_text(CHANGELOG)
    forbidden = [
        "Telegram",
        "Gemini",
        "GEMINI",
        "TELEGRAM",
        "secrets",
        "prompt",
        "sales-only",
        "playbook",
    ]
    found = [word for word in forbidden if word in text]
    assert_true(not found, f"Customer changelog contains internal words: {found}")


def find_header_row(ws, required_headers: set[str]) -> tuple[int, dict[str, int]]:
    for row_idx in range(1, min(ws.max_row, 20) + 1):
        values = [str(ws.cell(row=row_idx, column=col).value or "").strip() for col in range(1, ws.max_column + 1)]
        mapping = {value: idx + 1 for idx, value in enumerate(values) if value}
        if required_headers.issubset(mapping):
            return row_idx, mapping
    raise AssertionError(f"Required headers not found: {sorted(required_headers)}")


def check_portfolio() -> None:
    wb = load_workbook(PORTFOLIO, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header_row, headers = find_header_row(ws, {"Vendor", "Distributor KZ", "Distributor Status"})
    vendor_col = headers["Vendor"]
    distributor_col = headers["Distributor KZ"]
    status_col = headers["Distributor Status"]

    verified = 0
    non_verified = 0
    for row_idx in range(header_row + 1, ws.max_row + 1):
        vendor = str(ws.cell(row=row_idx, column=vendor_col).value or "").strip()
        distributor = str(ws.cell(row=row_idx, column=distributor_col).value or "").strip()
        status = str(ws.cell(row=row_idx, column=status_col).value or "").strip().lower()
        if not vendor or not distributor:
            continue
        if status in {"подтверждено", "проверенный", "verified", "confirmed"}:
            verified += 1
        else:
            non_verified += 1

    assert_true(verified > 0, "No verified distributor rows found")
    assert_true(non_verified > 0, "No non-verified distributor rows found; filter cannot be validated")


def check_static_hooks() -> None:
    text = read_text(APP)
    assert_true("def build_ai_first_sales_opportunities" in text, "AI-first sales helper is missing")
    assert_true("last_report_risk_sources" in text and '"vendors": item.get("vendors", [])' in text, "Risk sources do not preserve vendors")
    assert_true('"area": item.get("_ai_area", "ИТ/ИБ")' in text, "Risk sources do not preserve IT/IB area")
    assert_true("build_ai_first_sales_opportunities(risk_sources)" in text, "Sales playbook does not use AI-first opportunities")
    assert_true("def it_context_summary" in text and "ИТ-контекст" in text, "Client report IT context is missing")


def main() -> None:
    checks = [
        check_compile,
        check_version,
        check_customer_changelog,
        check_portfolio,
        check_static_hooks,
    ]
    for check in checks:
        check()
        print(f"OK {check.__name__}")
    print("SMOKE TEST PASSED")


if __name__ == "__main__":
    main()
