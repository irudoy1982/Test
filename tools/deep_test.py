from __future__ import annotations

import ast
import os
import re
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
APP = ROOT / "audit_app.py"


def assert_true(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def extract_function_source(module_text: str, function_name: str) -> str:
    tree = ast.parse(module_text)
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name == function_name:
            source = ast.get_source_segment(module_text, node)
            if source:
                return source
    raise AssertionError(f"Function not found: {function_name}")


def load_ai_first_helper():
    module_text = APP.read_text(encoding="utf-8")
    namespace = {
        "re": re,
        "manufacturers_for_report_item": lambda item: "FallbackVendor",
        "portfolio_vendors_for_report_item": lambda item: "Imperva, F5" if "WAF" in str(item.get("vendors")) else ", ".join(item.get("vendors", [])),
    }
    for name in (
        "normalize_vendor_key",
        "result_contains_any",
        "sales_override_for_item",
        "build_ai_first_sales_opportunities",
    ):
        exec(extract_function_source(module_text, name), namespace)
    return namespace["build_ai_first_sales_opportunities"]


def load_portfolio_helpers():
    module_text = APP.read_text(encoding="utf-8")
    names = [
        "normalize_vendor_key",
        "clean_vendor_display_name",
        "load_detailed_solution_vendor_map",
        "normalize_portfolio_header",
        "split_portfolio_list",
        "load_verified_distributor_map",
        "load_solution_vendor_map",
        "manufacturers_for_report_item",
        "verified_distributors_for_vendors",
        "portfolio_vendors_for_report_item",
    ]
    namespace = {
        "DETAILED_VENDOR_MATRIX_FILE": str(ROOT / "vendor_matrix_detailed.xlsx"),
        "os": os,
        "pd": pd,
        "re": re,
    }
    for name in names:
        exec(extract_function_source(module_text, name), namespace)
    return namespace


def test_ai_first_sales_behavior() -> None:
    build = load_ai_first_helper()
    rows = build(
        [
            {
                "level": "Высокий",
                "risk": "Публичные web-сервисы требуют WAF",
                "description": "Есть личный кабинет и интернет-магазин.",
                "impact": "Риск атак на приложение и простоя клиентских сервисов.",
                "recommendation": "Провести экспресс-оценку web-периметра; включить WAF/CDN; настроить контроль блокировок.",
                "vendors": ["WAF"],
                "area": "ИБ",
                "source": "ИИ",
            },
            {
                "level": "Высокий",
                "risk": "Базовый риск должен быть проигнорирован",
                "recommendation": "Не должен попасть в AI-first лист.",
                "source": "Базовые правила",
            },
        ],
        {"NGFW": "FortiGate", "3.2. Frontend": "Cloudflare"},
        {"has_public_web": True},
    )

    assert_true(len(rows) == 1, f"Expected exactly one AI opportunity, got {len(rows)}")
    assert_true(rows[0]["priority"] == "P2", "WAF should be P2 for this sales playbook")
    assert_true(rows[0]["source"] == "ИИ", "AI opportunity source must stay visible in playbook")
    assert_true(rows[0]["vendors"] == "Cloudflare, Fortinet, F5, Imperva", "WAF should prefer existing Cloudflare/Fortinet stack")
    assert_true("web" in rows[0]["problem"].lower(), "Risk title should be preserved")


def test_sales_overrides_for_mfa_and_legacy_os() -> None:
    build = load_ai_first_helper()
    rows = build(
        [
            {
                "level": "Критический",
                "risk": "Устаревшие операционные системы на рабочих станциях",
                "recommendation": "Закрыть риск устаревших ОС",
                "vendors": ["EDR/XDR"],
                "source": "ИИ",
            },
            {
                "level": "Высокий",
                "risk": "Отсутствие многофакторной аутентификации MFA",
                "recommendation": "Включить MFA",
                "vendors": ["PAM"],
                "source": "ИИ",
            },
        ],
        {"NGFW": "FortiGate", "1.5.1. Почтовая система": "Microsoft 365"},
        {"users": 120, "servers": 22},
    )
    legacy = next(row for row in rows if "Устаревшие ОС" in row["problem"])
    mfa = next(row for row in rows if row["problem"].startswith("MFA"))
    assert_true("Microsoft" in legacy["vendors"], "Legacy OS should point to Microsoft migration")
    assert_true("CrowdStrike" not in legacy["vendors"] and "Trend Micro" not in legacy["vendors"], "Legacy OS should not be sold as EDR")
    assert_true("Fortinet" in mfa["vendors"] and "Microsoft" in mfa["vendors"], "MFA should use Fortinet/Microsoft path")
    assert_true("CyberArk" not in mfa["vendors"] and "Wallix" not in mfa["vendors"], "MFA should not be mapped to PAM vendors")


def test_portfolio_category_to_verified_distributor() -> None:
    helpers = load_portfolio_helpers()
    vendors = helpers["portfolio_vendors_for_report_item"]({"vendors": ["WAF"], "risk": "WAF"})
    distributors = helpers["verified_distributors_for_vendors"]("WAF")
    assert_true("Check Point" in vendors, f"WAF should resolve to portfolio vendor, got: {vendors}")
    assert_true("MONT TECH" in distributors or "MUK" in distributors, f"WAF distributor should be verified, got: {distributors}")


def test_sales_fallback_hook_order() -> None:
    text = APP.read_text(encoding="utf-8")
    internal = extract_function_source(text, "make_internal_sales_excel")
    ai_call = internal.find("build_ai_first_sales_opportunities(risk_sources, results, context)")
    fallback_call = internal.find("build_sales_opportunities(results, context, roadmap_items)")
    enrich_call = internal.find("ensure_sales_playbook_priorities")
    assert_true(ai_call >= 0, "AI-first sales call is missing")
    assert_true(fallback_call >= 0, "Fallback sales call is missing")
    assert_true(enrich_call >= 0, "Expert sales prioritization is missing")
    assert_true(ai_call < fallback_call, "AI-first call must be evaluated before fallback")


def test_customer_report_context() -> None:
    text = APP.read_text(encoding="utf-8")
    report_source = extract_function_source(text, "make_expert_excel")
    assert_true("it_context_summary(results, context)" in report_source, "Report does not compute IT context")
    assert_true("ИТ-контекст" in report_source, "Report passport does not display IT context")
    assert_true("Фокус эксплуатации" in report_source, "Report passport does not display operational focus")


def test_sales_sheet_navigation_layout() -> None:
    text = APP.read_text(encoding="utf-8")
    internal = extract_function_source(text, "make_internal_sales_excel")
    assert_true("A1:H1" in internal, "Sales sheet title should span all 8 columns")
    assert_true("A3:H3" in internal, "Sales sheet company strip should span all 8 columns")
    assert_true("A5:H5" in internal, "Sales sheet company header should span all 8 columns")
    assert_true("Навигация:" in internal, "Sales sheet navigation hint is missing")
    assert_true("ws.column_dimensions['H'].width = 16" in internal, "Source column should remain visible")


def main() -> None:
    tests = [
        test_ai_first_sales_behavior,
        test_sales_overrides_for_mfa_and_legacy_os,
        test_portfolio_category_to_verified_distributor,
        test_sales_fallback_hook_order,
        test_customer_report_context,
        test_sales_sheet_navigation_layout,
    ]
    for test in tests:
        test()
        print(f"OK {test.__name__}")
    print("DEEP TEST PASSED")


if __name__ == "__main__":
    main()
