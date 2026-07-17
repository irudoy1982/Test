from __future__ import annotations

import ast
import os
import re
import zipfile
from io import BytesIO
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
    def fake_portfolio_vendors_by_categories(categories, preferred=None, exclude=None, gap_text=None, limit=6):
        if "WAF" in categories:
            return "Check Point"
        if "IDS/IPS" in categories:
            return gap_text or "Нет категории IDS/IPS в матрице"
        if "MFA" in categories:
            return gap_text or "Нет корректного MFA-вендора в матрице"
        if "Operating Systems" in categories:
            return "Microsoft, Qualys, Tenable, Rapid7"
        return ", ".join(preferred or []) or (gap_text or "-")

    namespace["portfolio_vendors_by_categories"] = fake_portfolio_vendors_by_categories
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
        "risk_semantic_key",
        "normalize_report_vendor_values",
        "solution_categories_for_report_item",
        "portfolio_manufacturers_for_report_item",
        "manufacturers_for_report_item",
        "verified_distributors_for_vendors",
        "portfolio_vendors_for_report_item",
        "portfolio_vendors_by_categories",
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


def load_report_logic_helpers():
    module_text = APP.read_text(encoding="utf-8")
    namespace = {"re": re}
    for name in (
        "risk_source_label",
        "risk_semantic_key",
        "network_segmentation_evidence",
        "neutralize_company_scale_language",
        "sanitize_ai_audit_narrative",
        "professionalize_risk_item",
        "russian_count",
        "infrastructure_profile",
        "sales_account_guidance",
        "build_sales_conversation_pack",
    ):
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
    assert_true(rows[0]["vendors"] == "Check Point", "WAF should be selected from portfolio matrix")
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
    assert_true("Нет корректного MFA-вендора в матрице" in mfa["vendors"], "MFA should expose portfolio gap instead of wrong vendor")
    assert_true("CyberArk" not in mfa["vendors"] and "Wallix" not in mfa["vendors"], "MFA should not be mapped to PAM vendors")


def test_ids_ips_exposes_matrix_gap() -> None:
    build = load_ai_first_helper()
    rows = build(
        [
            {
                "level": "Высокий",
                "risk": "Недостаточная защита от сетевых атак IDS/IPS",
                "recommendation": "Внедрить IDS/IPS",
                "vendors": ["IDS/IPS"],
                "source": "ИИ",
            },
        ],
        {"NGFW": "FortiGate"},
        {"users": 120, "servers": 22},
    )
    assert_true(len(rows) == 1, f"Expected one IDS/IPS row, got {len(rows)}")
    assert_true(rows[0]["priority"] == "P3", "IDS/IPS should be an уточнение, not P1/P2")
    assert_true("Нет категории IDS/IPS в матрице" in rows[0]["vendors"], "IDS/IPS should expose matrix gap")


def test_portfolio_category_to_verified_distributor() -> None:
    helpers = load_portfolio_helpers()
    vendors = helpers["portfolio_vendors_for_report_item"]({"vendors": ["WAF"], "risk": "WAF"})
    distributors = helpers["verified_distributors_for_vendors"]("WAF")
    assert_true("Check Point" in vendors, f"WAF should resolve to portfolio vendor, got: {vendors}")
    assert_true("MONT TECH" in distributors or "MUK" in distributors, f"WAF distributor should be verified, got: {distributors}")


def test_database_security_vendors_and_distributors() -> None:
    helpers = load_portfolio_helpers()
    vendors = helpers["portfolio_vendors_by_categories"](
        ["DB Security"],
        preferred=["Garda Technologies (Гарда)", "Imperva"],
    )
    distributors = helpers["verified_distributors_for_vendors"](vendors)
    assert_true("Garda Technologies" in vendors, f"Garda should resolve for DB Security, got: {vendors}")
    assert_true("Imperva" in vendors, f"Imperva should resolve for DB Security, got: {vendors}")
    assert_true("Axoft Kazakhstan" in distributors, f"Garda distributor should resolve to Axoft, got: {distributors}")
    assert_true("Softprom" in distributors, f"Imperva distributor should resolve to Softprom, got: {distributors}")


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


def test_customer_report_separates_solutions_from_manufacturers() -> None:
    text = APP.read_text(encoding="utf-8")
    report_source = extract_function_source(text, "make_expert_excel")
    assert_true("solution_categories_for_report_item(item)" in report_source, "Report solutions row should show solution classes, not vendors")
    assert_true("portfolio_manufacturers_for_report_item(item)" in report_source, "Report manufacturers row should use portfolio matrix")

    helpers = load_portfolio_helpers()
    waf_item = {
        "risk": "Публичные web-сервисы требуют WAF",
        "description": "Есть интернет-магазин и личный кабинет.",
        "recommendation": "Включить WAF/CDN.",
        "vendors": ["F5", "Imperva"],
    }
    solution = helpers["solution_categories_for_report_item"](waf_item)
    manufacturers = helpers["portfolio_manufacturers_for_report_item"](waf_item)
    assert_true("WAF" in solution and "F5" not in solution and "Imperva" not in solution, f"Solutions should be classes, got: {solution}")
    assert_true("Check Point" in manufacturers, f"WAF manufacturer should be resolved from matrix, got: {manufacturers}")


def test_segmentation_never_maps_to_dlp() -> None:
    helpers = load_portfolio_helpers()
    item = {
        "risk": "Недостаточная сегментация сети и контроль доступа",
        "description": "Архитектура VLAN и ACL не раскрыта.",
        "impact": "Риск доступа к конфиденциальным данным.",
        "recommendation": "Проверить межсегментные правила NGFW.",
        "vendors": [],
    }
    key = helpers["risk_semantic_key"](item)
    solution = helpers["solution_categories_for_report_item"](item)
    manufacturers = helpers["portfolio_manufacturers_for_report_item"](item)
    assert_true(key == "segmentation", f"Segmentation must win over generic confidentiality text, got: {key}")
    assert_true("VLAN" in solution and "DLP" not in solution, f"Segmentation solution is incorrect: {solution}")
    assert_true("Zecurion" not in manufacturers and "Forcepoint" not in manufacturers, f"DLP vendors leaked into segmentation: {manufacturers}")


def test_ospf_is_not_segmentation_evidence() -> None:
    helpers = load_report_logic_helpers()
    results = {
        "1.2.3. Маршрутизация": "Статическая, OSPF",
        "NGFW": "Fortinet FortiGate",
        "NAC": "Нет",
        "ZTNA": "Нет",
    }
    item = {
        "risk": "Недостаточная сегментация сети и контроль доступа",
        "description": "Используется OSPF, NAC и ZTNA отсутствуют.",
        "impact": "Возможно боковое перемещение.",
        "recommendation": "Внедрить сегментацию.",
        "vendors": ["DLP"],
        "_source": "ИИ",
    }
    normalized = helpers["professionalize_risk_item"](item, results, {"users": 120, "servers": 22})
    assert_true(helpers["network_segmentation_evidence"](results) == "unknown", "OSPF must not count as segmentation evidence")
    assert_true(normalized["level"] == "LOW", "Unverified segmentation must be a validation item, not a high confirmed risk")
    assert_true("требует подтверждения" in normalized["risk"].lower(), f"Unexpected title: {normalized['risk']}")
    assert_true(normalized["vendors"] == [], "No product should be prescribed before segmentation is confirmed")


def test_customer_and_sales_language_avoids_size_labels() -> None:
    helpers = load_report_logic_helpers()
    profile_title, profile_text = helpers["infrastructure_profile"]({
        "users": 120,
        "servers": 22,
        "small_company": False,
        "medium_company": True,
        "large_company": False,
        "enterprise_company": False,
    })
    combined = f"{profile_title} {profile_text}".lower()
    assert_true("малая" not in combined and "средняя" not in combined and "крупная" not in combined, combined)

    pack = helpers["build_sales_conversation_pack"](
        {"Наименование компании": "Demo"},
        {"MFA": "Нет", "Patch Management": "Нет", "EDR": "Нет", "SIEM": "Нет", "Резервное копирование": "Veeam"},
        {"users": 120, "servers": 22, "small_company": False, "medium_company": True, "has_public_web": False, "has_development": False},
        [],
        [{
            "priority": "P1",
            "problem": "MFA для критичных доступов",
            "offer": "MFA-пилот",
            "trigger": "MFA не указана",
            "vendors": "Fortinet",
            "next_step": "Уточнить IdP",
        }],
    )
    assert_true(len(pack["call_script"]) >= 8, "Senior call scenario should contain full discovery flow")
    assert_true(all("goal" in item for item in pack["call_script"]), "Every call stage needs a goal")
    assert_true(all(len(item) == 3 for item in pack["questions"]), "Discovery questions need an explicit purpose")
    assert_true(all(len(item) == 3 for item in pack["objections"]), "Objections need a follow-up question")
    sales_text = str(pack).lower()
    assert_true("маленькая компания" not in sales_text and "средняя инфраструктура" not in sales_text, "Sales text contains harmful size labels")


def test_sales_sheet_navigation_layout() -> None:
    text = APP.read_text(encoding="utf-8")
    internal = extract_function_source(text, "make_internal_sales_excel")
    assert_true("A1:J1" in internal, "Sales sheet title should span all 10 columns")
    assert_true("A3:J3" in internal, "Sales sheet company strip should span all 10 columns")
    assert_true("A5:J5" in internal, "Sales sheet company header should span all 10 columns")
    assert_true("Навигация:" in internal, "Sales sheet navigation hint is missing")
    assert_true("Производители из портфеля" in internal, "Sales vendors column should be named as manufacturers")
    assert_true("Решения из портфеля" not in internal, "Sales vendors column should not be mislabeled as solutions")
    assert_true("Факт / что подтвердить" in internal, "Sales sheet must distinguish facts from hypotheses")
    assert_true("Ценность для клиента" in internal, "Sales sheet must explain client value")
    assert_true("Кого подключить" in internal, "Sales sheet must include stakeholder guidance")
    assert_true("ws.column_dimensions['J'].width = 16" in internal, "Source column should remain visible")


def test_presentation_template_rendering() -> None:
    module_text = APP.read_text(encoding="utf-8")
    namespace = {"BytesIO": BytesIO, "re": re}
    exec(extract_function_source(module_text, "render_audit_presentation_template"), namespace)
    render = namespace["render_audit_presentation_template"]

    for brand in ("khalil", "btg"):
        template = ROOT / "static" / f"audit_presentation_{brand}.pptx"
        with zipfile.ZipFile(template, "r") as archive:
            source_xml = "\n".join(
                archive.read(name).decode("utf-8")
                for name in archive.namelist()
                if name.startswith("ppt/slides/slide") and name.endswith(".xml")
            )
        tokens = set(re.findall(r"\{\{([A-Z0-9_]+)\}\}", source_xml))
        replacements = {token: f"Тест & проверка {token}" for token in tokens}
        rendered = render(template, replacements)

        with zipfile.ZipFile(BytesIO(rendered), "r") as archive:
            slide_names = [
                name for name in archive.namelist()
                if name.startswith("ppt/slides/slide") and name.endswith(".xml")
            ]
            rendered_xml = "\n".join(archive.read(name).decode("utf-8") for name in slide_names)
        assert_true(len(slide_names) == 8, f"{brand}: expected 8 slides, got {len(slide_names)}")
        assert_true("{{" not in rendered_xml, f"{brand}: unresolved presentation placeholders")
        assert_true("Тест &amp; проверка" in rendered_xml, f"{brand}: XML escaping failed")


def test_presentation_text_is_self_contained() -> None:
    module_text = APP.read_text(encoding="utf-8")
    namespace = {"re": re}
    exec(extract_function_source(module_text, "presentation_text"), namespace)
    clean = namespace["presentation_text"]
    assert_true(clean("Нормальный русский текст") == "Нормальный русский текст", "Normal text was changed")
    assert_true(clean("  Строка   с пробелами. ") == "Строка с пробелами", "Whitespace cleanup failed")
    shortened = clean("Очень длинная рекомендация " * 20, 80)
    assert_true(len(shortened) <= 80 and shortened.endswith("."), "Long text was not shortened to a complete phrase")


def main() -> None:
    tests = [
        test_ai_first_sales_behavior,
        test_sales_overrides_for_mfa_and_legacy_os,
        test_ids_ips_exposes_matrix_gap,
        test_portfolio_category_to_verified_distributor,
        test_database_security_vendors_and_distributors,
        test_sales_fallback_hook_order,
        test_customer_report_context,
        test_customer_report_separates_solutions_from_manufacturers,
        test_segmentation_never_maps_to_dlp,
        test_ospf_is_not_segmentation_evidence,
        test_customer_and_sales_language_avoids_size_labels,
        test_sales_sheet_navigation_layout,
        test_presentation_template_rendering,
        test_presentation_text_is_self_contained,
    ]
    for test in tests:
        test()
        print(f"OK {test.__name__}")
    print("DEEP TEST PASSED")


if __name__ == "__main__":
    main()
