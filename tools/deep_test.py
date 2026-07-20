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

    change_item = {
        "risk": "Отсутствие формализованного управления изменениями и конфигурациями",
        "impact": "Ошибки могут привести к уязвимостям.",
    }
    change_solution = helpers["solution_categories_for_report_item"](change_item)
    change_vendors = helpers["portfolio_manufacturers_for_report_item"](change_item)
    assert_true("Change Management" in change_solution, f"Unexpected change-management solution: {change_solution}")
    assert_true("Qualys" not in change_vendors and "Tenable" not in change_vendors, f"VM vendors leaked into change management: {change_vendors}")

    legacy_vendors = helpers["portfolio_manufacturers_for_report_item"]({"risk": "Устаревшие Windows 7 на рабочих станциях"})
    assert_true(legacy_vendors == "Microsoft", f"Legacy OS slide should recommend Microsoft only, got: {legacy_vendors}")


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


def test_wifi_and_dr_semantics_are_stable() -> None:
    helpers = load_portfolio_helpers()
    wifi_item = {
        "risk": "Перегрузка Wi‑Fi без централизованного управления",
        "description": "12 точек доступа обслуживают 500–600 одновременных подключений.",
        "recommendation": "Провести радиообследование и пилот WLAN-контроллера.",
        "vendors": [],
    }
    assert_true(
        helpers["risk_semantic_key"](wifi_item) == "wifi_capacity",
        "Non-breaking Wi-Fi hyphen must not turn a WLAN gap into IT monitoring",
    )
    wifi_solution = helpers["solution_categories_for_report_item"](wifi_item)
    wifi_vendors = helpers["portfolio_manufacturers_for_report_item"](wifi_item)
    assert_true("WLAN" in wifi_solution and "Wi-Fi" in wifi_solution, f"Incorrect Wi-Fi solution: {wifi_solution}")
    assert_true(
        "Cisco" in wifi_vendors and "Huawei" in wifi_vendors,
        f"Portfolio Wi-Fi manufacturers are missing: {wifi_vendors}",
    )

    dr_item = {
        "risk": "Не описан план аварийного восстановления критичных ИТ-сервисов",
        "description": "Нужно восстановить ERP, CRM и почту в пределах RTO/RPO.",
        "recommendation": "Провести DR-учение и оформить runbook.",
    }
    assert_true(
        helpers["risk_semantic_key"](dr_item) == "dr",
        "A DR recommendation mentioning mail must not map to Mail Security",
    )


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

    brand_identity = {
        "khalil": ("ТОО «Khalil Trade»", "info@khalilgroup.kz", "+7 706 701 48 35", "2020", "Bolashak Tamer Group"),
        "btg": ("ТОО «Bolashak Tamer Group»", "info@btgroup.kz", "+7 706 700 48 35", "2019", "Khalil Trade"),
    }
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
        assert_true(len(slide_names) == 13, f"{brand}: expected 13 slides, got {len(slide_names)}")
        assert_true("Команда для реализации изменений" in rendered_xml, f"{brand}: company profile slide is missing")
        company_name, email, phone, founded_year, foreign_brand = brand_identity[brand]
        assert_true(company_name in rendered_xml, f"{brand}: company name is missing")
        assert_true(email in rendered_xml and phone in rendered_xml, f"{brand}: contact details are missing")
        assert_true(founded_year in rendered_xml, f"{brand}: founding year is missing")
        assert_true(foreign_brand not in rendered_xml, f"{brand}: foreign brand data leaked into presentation")
        assert_true("{{" not in rendered_xml, f"{brand}: unresolved presentation placeholders")
        assert_true("C10001" not in rendered_xml and "C10002" not in rendered_xml, f"{brand}: maturity colors were not rendered")
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


def test_presentation_actions_are_complete_and_deduplicated() -> None:
    module_text = APP.read_text(encoding="utf-8")
    namespace = {"re": re, "expand_regulatory_references": lambda value: value}
    for name in (
        "presentation_text",
        "presentation_action_text",
        "risk_level_label",
        "risk_semantic_key",
        "presentation_recommendation_key",
        "presentation_presales_profile",
        "presentation_severity_style",
        "presentation_risk_entry",
    ):
        exec(extract_function_source(module_text, name), namespace)

    clean_action = namespace["presentation_action_text"]
    action = clean_action(
        "1. Провести аудит текущей конфигурации маршрутизации и сетевой топологии. "
        "2. Подготовить целевую архитектуру и план модернизации. "
        "3. Согласовать этапы внедрения.",
        120,
    )
    assert_true(action.endswith("."), f"Presentation action is incomplete: {action}")
    assert_true(not re.search(r"\b\d+\.\s*$", action), f"Presentation action ends with a dangling number: {action}")

    recommendation_key = namespace["presentation_recommendation_key"]
    generic_mfa = recommendation_key({"domain": "ИБ", "action": "Внедрить MFA для критичных и удаленных доступов."})
    explicit_mfa = recommendation_key({"risk": "Отсутствует многофакторная аутентификация"})
    assert_true(generic_mfa == explicit_mfa == "mfa", "MFA recommendations are not deduplicated semantically")
    software_lifecycle = recommendation_key({
        "risk": "Проблемы с управлением жизненным циклом программного обеспечения",
        "recommendation": "Вести реестр версий и обновлений.",
    })
    assert_true(software_lifecycle == "itam", "Software lifecycle recommendation should map to ITAM, not AppSec or patching")

    change_key = recommendation_key({
        "risk": "Отсутствие формализованного управления изменениями и конфигурациями",
        "impact": "Ошибки могут привести к уязвимостям.",
    })
    assert_true(change_key == "change_management", "Change management should not map to vulnerability management")

    _, monitoring_profile = namespace["presentation_presales_profile"]({
        "risk": "Отсутствие централизованного мониторинга производительности",
        "recommendation": "Внедрить Zabbix или Prometheus.",
    })
    profile_text = str(monitoring_profile)
    assert_true("Zabbix" not in profile_text and "Prometheus" not in profile_text, "Non-portfolio monitoring brands leaked into presentation")
    assert_true(monitoring_profile["action"].endswith("."), "Monitoring action must be a complete sentence")

    network_risk = namespace["presentation_risk_entry"]({
        "level": "HIGH",
        "risk": "Недостаточная производительность и масштабируемость сетевой инфраструктуры",
        "impact": "Масштабирование затруднено, сегментация неизвестна.",
        "recommendation": "Провести аудит сети.",
    })
    assert_true("сегментац" not in network_risk["impact"].lower(), "Network performance slide must not invent missing segmentation")
    nac_risk = namespace["presentation_risk_entry"]({
        "_source": "Groq",
        "level": "MEDIUM",
        "risk": "Отсутствие NAC приводит к неавтоматизированному контролю доступа устройств",
        "impact": "Увеличение вероятности lateral movement и компрометации критических серверов.",
        "recommendation": "Внедрить NAC.",
    })
    assert_true(
        nac_risk["title"] == "Допуск устройств к сети не контролируется автоматически"
        and "lateral movement" not in nac_risk["impact"].lower(),
        "Known NAC findings must use the fact-safe presales title and impact",
    )
    assert_true(len(nac_risk["title"]) <= 58, "Risk-card title can overlap the impact block")
    complete_title = namespace["presentation_risk_entry"]({
        "level": "HIGH",
        "risk": "Не описан план аварийного восстановления критичных сервисов",
        "impact": "Восстановление может не уложиться в согласованное время.",
        "recommendation": "Определить RTO/RPO и провести учение.",
    })
    assert_true(
        complete_title["title"].endswith("сервисов"),
        f"Risk title was cut mid-sentence: {complete_title['title']}",
    )
    dlp_risk = namespace["presentation_risk_entry"]({
        "_source": "Groq",
        "level": "HIGH",
        "risk": "Отсутствие DLP повышает вероятность утечки персональных",
        "impact": "Возможны регуляторные последствия.",
        "recommendation": "Провести пилот DLP.",
    })
    assert_true(
        dlp_risk["title"] == "Отсутствие DLP повышает риск утечки персональных данных",
        "DLP risk title must be complete and use the fact-safe presales profile",
    )


def test_it_maturity_measures_controls_not_infrastructure_size() -> None:
    module_text = APP.read_text(encoding="utf-8")
    namespace: dict[str, object] = {"re": re}
    exec(extract_function_source(module_text, "calculate_it_maturity_score"), namespace)
    score = namespace["calculate_it_maturity_score"](
        420, ["Windows 10", "Windows 11"], 420,
        True, 1000, 200, ["OSPF"], 14, "FortiGate",
        True, 4, 38, ["VMware vSphere"], "Veeam",
        True, ["HDD", "SSD"], 24, 16, ["RAID 10"],
        True, True, True, 12, ["Python"], True,
    )
    assert_true(score <= 90, f"Questionnaire-only IT maturity cannot imply full optimization, got {score}")

    degraded_score = namespace["calculate_it_maturity_score"](
        650, ["Windows 11", "macOS"], 650,
        True, 1000, 50, ["OSPF"], 12, "Check Point",
        True, 8, 60, ["VMware"], "Veeam",
        True, ["HDD", "SSD"], 36, 12, ["RAID 6", "RAID 10"],
        True, True, False, 0, [], False,
        wifi_enabled=True,
        wifi_ctrl_enabled=False,
        operational_notes=[
            "CMDB отсутствует; изменения согласуются в чатах.",
            "Wi-Fi без единой панели; capacity planning не формализован.",
            "RTO и RPO не согласованы, восстановление тестируется нерегулярно.",
            "СХД заполнена на 84%.",
        ],
    )
    assert_true(
        degraded_score < 60,
        f"Confirmed operational IT gaps must materially reduce maturity, got {degraded_score}",
    )


def test_security_maturity_is_normalized_and_evidence_capped() -> None:
    module_text = APP.read_text(encoding="utf-8")
    namespace: dict[str, object] = {}
    exec(extract_function_source(module_text, "calculate_weighted_security_score"), namespace)
    controls = [(True, f"Vendor {index}", weight) for index, weight in enumerate((8, 12, 8, 10, 6, 5), start=1)]
    score = namespace["calculate_weighted_security_score"](True, controls)
    assert_true(score == 92, f"Self-reported security maturity must be capped at 92%, got {score}")

    partial = [(True, "Vendor", 10), (False, "", 10), (False, "", 10)]
    partial_score = namespace["calculate_weighted_security_score"](True, partial)
    assert_true(partial_score == 31, f"Security score must be normalized by available weight, got {partial_score}")


def test_confirmed_it_gaps_must_be_covered_by_ai() -> None:
    module_text = APP.read_text(encoding="utf-8")
    labels = {
        key: key
        for key in (
            "wifi_capacity", "network_performance", "virtualization", "storage", "it_monitoring",
            "itam", "change_management", "dr",
        )
    }
    namespace = {
        "re": re,
        "IT_GAP_LABELS": labels,
    }
    exec(extract_function_source(module_text, "confirmed_it_gap_topics"), namespace)
    results = {
        "_user_count": 650,
        "WiFi Точки": 12,
        "_main_speed": "1000 Mbit/s",
        "_back_speed": "50 Mbit/s",
        "WiFi Контроллер": "Нет",
        "1.1. Примечание": "Единый CMDB отсутствует.",
        "1.2. Примечание": "Wi-Fi перегружен, резервный канал слабый, мониторинг без единой панели.",
        "1.3. Примечание": "Capacity planning не формализован; RTO/RPO не согласованы, восстановление тестируется нерегулярно.",
        "1.4. Примечание": "СХД заполнена на 84%, прогноз исчерпания не утвержден.",
        "1.5. Примечание": "Изменения согласуются в чатах, календарь изменений отсутствует.",
    }
    expected = namespace["confirmed_it_gap_topics"](results)
    assert_true(
        set(expected) == set(labels),
        f"Not all explicit IT gaps were detected: {sorted(expected)}",
    )

    namespace["risk_semantic_key"] = None
    exec(extract_function_source(module_text, "risk_semantic_key"), namespace)
    exec(extract_function_source(module_text, "ai_it_gap_coverage"), namespace)
    ai_items = [
        {"risk": "Перегруженная Wi-Fi сеть без централизованного WLAN-контроллера"},
        {"risk": "Слабый резервный канал не обеспечивает отказоустойчивость WAN"},
        {"risk": "Capacity planning виртуализации не формализован"},
        {"risk": "СХД требует контроля емкости и производительности"},
        {"risk": "Эксплуатационный мониторинг работает без единой панели"},
        {"risk": "CMDB и учет активов требуют централизации"},
    ]
    matched, missing = namespace["ai_it_gap_coverage"](ai_items, expected)
    assert_true(len(matched) == 6, f"Expected six covered IT gaps, got {matched}")
    assert_true(set(missing) == {"change_management", "dr"}, f"Unexpected missing IT gaps: {missing}")

    channel_only_matched, channel_only_missing = namespace["ai_it_gap_coverage"](
        [{"risk": "Резервный канал требует увеличения пропускной способности и проверки failover"}],
        expected,
    )
    assert_true("network_performance" in channel_only_matched, "WAN recommendation must cover network resilience")
    assert_true("wifi_capacity" in channel_only_missing, "WAN recommendation must not hide a confirmed Wi-Fi gap")

    combined_item = [{
        "risk": "CMDB и управление изменениями требуют формализации",
        "description": "Учет активов и планы отката не связаны.",
        "recommendation": "Настроить CMDB, согласование изменений и контроль восстановления по RTO/RPO.",
    }]
    combined_matched, _ = namespace["ai_it_gap_coverage"](combined_item, expected)
    assert_true(
        {"itam", "change_management", "dr"}.issubset(combined_matched),
        f"One complete AI recommendation must be able to cover related gaps: {combined_matched}",
    )


def test_ai_presentation_recommendation_path() -> None:
    module_text = APP.read_text(encoding="utf-8")
    namespace = {
        "expand_regulatory_references": lambda value: value,
        "presentation_presales_profile": lambda item: ("nac", {}),
        "presentation_text": lambda value, limit=None: str(value),
        "presentation_action_text": lambda value, limit=None: str(value),
        "solution_categories_for_report_item": lambda item: "NAC",
        "portfolio_manufacturers_for_report_item": lambda item: "Fortinet",
        "split_portfolio_list": lambda value: [part.strip() for part in str(value).split(",") if part.strip()],
        "presentation_evidence_for_key": lambda *args: "NAC в анкете: Нет",
        "REGULATORY_CATALOG": {},
        "presentation_legal_basis": lambda *args: "Применимость подтверждена",
        "presentation_severity_style": lambda level: ("MEDIUM", "#F4B400", "#1F2937"),
        "risk_level_label": lambda level: "Средний",
        "presentation_success_metric": lambda key: "Все подключения идентифицируются",
    }
    exec(extract_function_source(module_text, "presentation_recommendation_entry"), namespace)
    rendered = namespace["presentation_recommendation_entry"]({
        "_source": "Groq",
        "level": "MEDIUM",
        "risk": "Контроль допуска устройств к сети",
        "recommendation": "Провести пилот NAC.",
        "evidence": ["NAC в анкете: Нет"],
    })
    assert_true(rendered["key"] == "nac", "AI recommendation lost its semantic key")
    assert_true(rendered["action"] == "Провести пилот NAC.", "AI recommendation path did not render")


def test_presentation_evidence_and_maturity_palette() -> None:
    module_text = APP.read_text(encoding="utf-8")
    namespace = {"re": re}
    for name in ("presentation_text", "presentation_action_text", "presentation_maturity_style", "presentation_evidence_for_key"):
        exec(extract_function_source(module_text, name), namespace)

    evidence = namespace["presentation_evidence_for_key"](
        "itam",
        {"MFA": "Нет"},
        {"users": 120, "servers": 22},
        {"description": "Компрометация учетной записи"},
    )
    assert_true("компрометац" not in evidence.lower(), f"ITAM evidence is semantically wrong: {evidence}")
    assert_true("реестр" in evidence.lower(), f"ITAM evidence must state the actual data gap: {evidence}")

    wifi_evidence = namespace["presentation_evidence_for_key"](
        "wifi_capacity",
        {
            "_user_count": 650,
            "Wi-Fi Точки доступа": 12,
            "Wi-Fi Контроллер": "Нет",
        },
        {"users": 650, "servers": 22},
        {},
    )
    assert_true(
        "650" in wifi_evidence and "12" in wifi_evidence and "Нет" in wifi_evidence,
        f"Wi-Fi evidence lost questionnaire values: {wifi_evidence}",
    )

    palette = namespace["presentation_maturity_style"]
    assert_true(palette(20)[0] == "#D92D20", "Low maturity must be red")
    assert_true(palette(55)[0] == "#F4B400", "Mid maturity must be yellow")
    assert_true(palette(85)[0] == "#13877C", "High maturity must be green")


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
        test_wifi_and_dr_semantics_are_stable,
        test_ospf_is_not_segmentation_evidence,
        test_customer_and_sales_language_avoids_size_labels,
        test_sales_sheet_navigation_layout,
        test_presentation_template_rendering,
        test_presentation_text_is_self_contained,
        test_presentation_actions_are_complete_and_deduplicated,
        test_ai_presentation_recommendation_path,
        test_it_maturity_measures_controls_not_infrastructure_size,
        test_security_maturity_is_normalized_and_evidence_capped,
        test_confirmed_it_gaps_must_be_covered_by_ai,
        test_presentation_evidence_and_maturity_palette,
    ]
    for test in tests:
        test()
        print(f"OK {test.__name__}")
    print("DEEP TEST PASSED")


if __name__ == "__main__":
    main()
