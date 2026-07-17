from __future__ import annotations

import py_compile
import hashlib
import re
import zipfile
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
APP = ROOT / "audit_app.py"
CHANGELOG = ROOT / "CHANGELOG_CUSTOMER.md"
PORTFOLIO = ROOT / "vendor_matrix_detailed.xlsx"
PRESENTATION_TEMPLATES = [
    ROOT / "static" / "audit_presentation_khalil.pptx",
    ROOT / "static" / "audit_presentation_btg.pptx",
]
PRESENTATION_QR_ASSETS = [
    ROOT / "static" / "presentation_khalil_qr.png",
    ROOT / "static" / "presentation_btg_qr.png",
]


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
    assert_true(match.group(1) == "12.8-dev", f"Unexpected APP_VERSION: {match.group(1)}")


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
    assert_true("build_ai_first_sales_opportunities(risk_sources, results, context)" in text, "Sales playbook does not use AI-first opportunities")
    assert_true("ensure_sales_playbook_priorities" in text, "Expert sales prioritization is missing")
    assert_true("def it_context_summary" in text and "ИТ-контекст" in text, "Client report IT context is missing")
    assert_true("def make_audit_presentation" in text, "Presentation generator is missing")
    assert_true("cached_presentation_bytes" in text, "Presentation download state is missing")
    assert_true("Скачать экспертный отчет (XLSX)" not in text, "Client XLSX download must stay hidden")
    assert_true("Скачать заключение по аудиту" in text, "Client presentation download is missing")
    assert_true("Скачать презентацию аудита (PPTX)" not in text, "Technical PPTX label must stay hidden")
    assert_true("st-key-presentation_download" in text, "Presentation download styling hook is missing")
    assert_true('"suffix": ".pptx"' in text and "Audit_Presentation_" in text, "Telegram presentation attachment is missing")


def check_presentation_templates() -> None:
    required = {
        "{{COMPANY}}",
        "{{IT_SCORE}}",
        "{{SUMMARY_1}}",
        "{{RISK_1_TITLE}}",
        "{{RISK_1_RECOMMENDATION}}",
        "{{FOCUS_1_TITLE}}",
        "{{FOCUS_1_TEXT}}",
        "{{REC_1_TITLE}}",
        "{{REC_1_ACTION}}",
        "{{REC_1_SOLUTION}}",
        "{{REC_1_VENDORS}}",
        "{{REC_8_TITLE}}",
        "{{REC_8_ACTION}}",
        "{{ROADMAP_1_1}}",
        "{{DECISION_1}}",
    }
    for template, qr_asset in zip(PRESENTATION_TEMPLATES, PRESENTATION_QR_ASSETS):
        assert_true(template.exists(), f"Presentation template is missing: {template.name}")
        with zipfile.ZipFile(template, "r") as archive:
            slide_xml = "\n".join(
                archive.read(name).decode("utf-8")
                for name in archive.namelist()
                if name.startswith("ppt/slides/slide") and name.endswith(".xml")
            )
            media_hashes = {
                hashlib.sha256(archive.read(name)).hexdigest()
                for name in archive.namelist()
                if name.startswith("ppt/media/")
            }
        missing = sorted(token for token in required if token not in slide_xml)
        assert_true(not missing, f"{template.name} is missing placeholders: {missing}")
        qr_hash = hashlib.sha256(qr_asset.read_bytes()).hexdigest()
        assert_true(qr_hash in media_hashes, f"{template.name} does not embed its contact QR")
    for qr_asset in PRESENTATION_QR_ASSETS:
        assert_true(qr_asset.exists(), f"Presentation QR is missing: {qr_asset.name}")
        assert_true(qr_asset.stat().st_size > 2000, f"Presentation QR is unexpectedly small: {qr_asset.name}")


def main() -> None:
    checks = [
        check_compile,
        check_version,
        check_customer_changelog,
        check_portfolio,
        check_static_hooks,
        check_presentation_templates,
    ]
    for check in checks:
        check()
        print(f"OK {check.__name__}")
    print("SMOKE TEST PASSED")


if __name__ == "__main__":
    main()
