import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import os
import html
import base64
import json
import zlib
import threading
import time
import random
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

#----------ИИ-----------
# --- AI BLOCK START ---

def sanitize_for_ai(c_info, results):
    forbidden = [
        "Наименование компании",
        "Сайт компании",
        "Email",
        "ФИО контактного лица",
        "Должность",
        "Контактный телефон"
    ]
    safe_client = {k: v for k, v in c_info.items() if k not in forbidden}
    safe_results = {
        k: v for k, v in results.items()
        if not any(f.lower() in str(k).lower() for f in forbidden)
    }
    return safe_client, safe_results

def load_vendor_matrix():
    try:
        df = pd.read_excel("Портфель для отчета.xlsx")

        vendors_text = ""

        for _, row in df.iterrows():
            row_text = " | ".join(
                [str(x) for x in row.values if pd.notna(x)]
            )
            vendors_text += row_text + "\n"

        return vendors_text

    except Exception as e:
        return f"Ошибка загрузки вендоров: {e}"


def get_app_secret(name, default=None):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


def get_regulators_by_industry(industry):
    regulators = {
        "Финтех / Банки": """
- Национальный Банк РК
- PCI DSS
- ISO 27001
- Постановления НБРК по ИБ
""",

        "Госсектор": """
- ГОСТ РК 34
- Требования ГТС
- Требования ИБ государственных ИС
- ISO 27001
""",

        "Ритейл / E-commerce": """
- PCI DSS
- Закон РК о персональных данных
- ISO 27001
""",

        "IT / Разработка": """
- OWASP ASVS
- Secure SDLC
- ISO 27001
- SOC2
""",

        "Производство": """
- ISA/IEC 62443
- ISO 27001
- Требования по защите АСУ ТП
"""
    }

    return regulators.get(
        industry,
        """
- ISO 27001
- Закон РК о персональных данных
"""
    )

def generate_rule_based_risks(results, context):

    risks = []
    users = context.get("users", 0)

    has_critical_systems = context.get(
        "has_critical_systems",
        False
    )

    has_personal_data = context.get(
        "has_personal_data",
        False
    )

    has_public_web = context.get(
        "has_public_web",
        False
    )

    has_development = context.get(
        "has_development",
        False
    )

    large_company = context.get(
        "large_company",
        False
    )

    enterprise_company = context.get(
        "enterprise_company",
        False
    )

    users = results.get("_user_count", 0)

    # =========================
    # VPN WITHOUT MFA
    # =========================

    if results.get("VPN") != "Нет" and results.get("MFA") == "Нет":

        risks.append({
            "level": "CRITICAL",
            "risk": "Удаленный доступ без MFA",
            "description": "VPN-доступ реализован без многофакторной аутентификации.",
            "impact": "Высокий риск компрометации учетных записей и несанкционированного доступа.",
            "recommendation": "Внедрить MFA для VPN, административного доступа и критичных систем.",
            "regulators": ["ISO 27001", "NIST", "PCI DSS"],
            "vendors": ["Cisco Duo", "Microsoft Entra ID", "FortiAuthenticator"]
        })

        # =========================
    # NO SIEM
    # =========================

    if results.get("Блок 2. SIEM") == "Нет":

        severity = "MEDIUM"

        if (
            has_critical_systems
            or has_personal_data
            or large_company
            or enterprise_company
        ):
            severity = "HIGH"

        risks.append({
            "level": severity,
            "risk": "Отсутствует централизованный мониторинг ИБ",
            "description": (
                "События безопасности не агрегируются "
                "в единой системе мониторинга."
            ),
            "impact": (
                "Увеличение времени обнаружения атак "
                "и расследования инцидентов."
            ),
            "recommendation": (
                "Рассмотреть внедрение SIEM "
                "или подключение внешнего SOC."
            ),
            "regulators": [
                "ISO 27001",
                "NIST CSF"
            ],
            "vendors": [
                "Microsoft Sentinel",
                "IBM QRadar",
                "Splunk"
            ]
        })

    # =========================
    # EPP WITHOUT EDR
    # =========================

    if results.get("Антивирус") != "Нет" and results.get("EDR") == "Нет":

        risks.append({
            "level": "HIGH",
            "risk": "Endpoint-защита ограничена только антивирусом",
            "description": "Используется только базовая антивирусная защита без EDR/XDR-функциональности.",
            "impact": "Низкая эффективность обнаружения сложных атак, ransomware и lateral movement.",
            "recommendation": "Рассмотреть внедрение EDR/XDR платформы.",
            "regulators": ["NIST", "MITRE ATT&CK"],
            "vendors": ["Defender for Endpoint", "CrowdStrike", "SentinelOne"]
        })

    # =========================
    # LARGE INFRA WITHOUT SEGMENTATION
    # =========================

    if users > 100 and results.get("Сегментация сети") == "Нет":

        risks.append({
            "level": "HIGH",
            "risk": "Отсутствует сегментация сети",
            "description": "Крупная инфраструктура эксплуатируется без сетевой сегментации.",
            "impact": "Высокий риск lateral movement и распространения malware между сегментами.",
            "recommendation": "Реализовать VLAN/ACL/Zero Trust сегментацию.",
            "regulators": ["ISO 27001", "NIST"],
            "vendors": ["Cisco", "Aruba", "Fortinet"]
        })

    # =========================
    # BACKUP RISKS
    # =========================

    if results.get("Резервное копирование") != "Нет" and results.get("Immutable Backup") == "Нет":

        risks.append({
            "level": "HIGH",
            "risk": "Backup не защищен от ransomware",
            "description": "Отсутствует immutable/offline backup.",
            "impact": "Риск уничтожения резервных копий при ransomware-инциденте.",
            "recommendation": "Внедрить immutable backup и air-gap копии.",
            "regulators": ["NIST", "ISO 27001"],
            "vendors": ["Veeam", "Commvault", "Rubrik"]
        })

    # =========================
    # PATCH MANAGEMENT
    # =========================

    old_os_detected = any([
        results.get("ОС АРМ (Windows XP/Vista/7/8)", 0) > 0,
        results.get("ОС Сервера (Windows Server 2008/2012 R2)", 0) > 0
    ])

    if results.get("Patch Management") == "Нет":

        severity = "MEDIUM"

        if old_os_detected:
            severity = "CRITICAL"

        elif users > 100:
            severity = "HIGH"

        risks.append({
            "level": severity,
            "risk": "Отсутствует централизованный Patch Management",
            "description": (
                "Обновления устанавливаются "
                "несистемно либо вручную."
            ),
            "impact": (
                "Высокая вероятность эксплуатации "
                "известных уязвимостей."
            ),
            "recommendation": (
                "Внедрить централизованное "
                "управление обновлениями."
            ),
            "regulators": [
                "CIS Controls",
                "NIST"
            ],
            "vendors": [
                "ManageEngine",
                "Ivanti",
                "Tanium",
                "Microsoft MECM"
            ]
        })
    # =========================
    # LEGACY OS
    # =========================

    legacy_workstations = (
        results.get(
            "ОС АРМ (Windows XP/Vista/7/8)",
            0
        )
    )

    legacy_servers = (
        results.get(
            "ОС Сервера (Windows Server 2008/2012 R2)",
            0
        )
    )

    if legacy_workstations > 0 or legacy_servers > 0:

        risks.append({

            "level": "CRITICAL",

            "risk": "Использование устаревших операционных систем",

            "description":
                f"Обнаружены устаревшие ОС. "
                f"АРМ: {legacy_workstations}, "
                f"Серверы: {legacy_servers}.",

            "impact":
                "Уязвимости больше не исправляются "
                "производителем.",

            "recommendation":
                "Разработать программу миграции "
                "на поддерживаемые версии ОС.",

            "regulators": [
                "ISO 27001",
                "CIS Controls"
            ],

            "vendors": [
                "Microsoft",
                "Red Hat",
                "VMware",
                "Citrix"
            ]
        })
    return risks
def ai_generate_risks_and_recs(c_info, results):
    import google.generativeai as genai
    import json
    import streamlit as st

    api_key = get_app_secret("GEMINI_API_KEY")

    if not api_key:
        return []

    try:
        genai.configure(api_key=api_key)

        available_models = [
            m.name for m in genai.list_models()
            if 'generateContent' in m.supported_generation_methods
        ]

        model_name = (
            available_models[0]
            if available_models
            else 'gemini-1.5-flash'
        )

        model = genai.GenerativeModel(model_name)

        safe_client, safe_results = sanitize_for_ai(
            c_info,
            results
        )

        vendor_context = load_vendor_matrix()

        regulator_context = get_regulators_by_industry(
            c_info.get("Сфера деятельности", "")
        )

        prompt = f"""
Выступай как аудитор Big4, CISO, CTO, Enterprise Security Architect и Lead Presale Engineer.

ЗАДАЧА:
Провести профессиональный анализ ИТ и ИБ инфраструктуры компании на основании опросника.

ВАЖНЫЕ ПРАВИЛА:

1. Не придумывай технологии без оснований в данных.

2. Не создавай риски только потому, что отсутствует продукт.

3. Оценивай бизнес-контекст компании.

4. Для каждой рекомендации сначала определяй КАТЕГОРИЮ решения:

Примеры:

Отсутствует SIEM
→ категория SIEM

Отсутствует Patch Management
→ категория Patch Management

Отсутствует Backup
→ категория Backup

Отсутствует CASB
→ категория CASB

Отсутствует ZTNA
→ категория ZTNA

Отсутствует SAST
→ категория Application Security

Устаревшие ОС
→ категория Migration Project

Устаревшие серверы
→ категория Infrastructure Modernization

5. Для категории решения сначала используй вендоров из каталога.

6. Если в каталоге отсутствуют подходящие решения,
разрешается использовать мировых лидеров рынка.

7. Никогда не предлагай:
- антивирусы
- EDR
- NGFW
- DLP

для устранения проблем устаревших операционных систем.

8. Для Legacy OS основная рекомендация:

- миграция
- модернизация
- виртуализация
- сегментация
- изоляция

9. Не предлагай SOAR как критическую необходимость
для компаний без зрелого SOC.

10. Не предлагай ZTNA как высокий риск без признаков удаленного доступа.

11. Для каждого риска указывай не более 3 наиболее релевантных решений.

12. Используй вендоров из каталога как приоритетных.

13. Если подходящего решения в каталоге нет,
подбери наиболее распространенные мировые решения.

14. Не дублируй риски.

15. Максимум 8 наиболее важных рисков.

16. Приоритет риска определяй по шкале:

CRITICAL
HIGH
MEDIUM
LOW

17. Особое внимание уделяй:

- устаревшим ОС
- отсутствию резервного копирования
- отсутствию MFA
- отсутствию Patch Management
- отсутствию сегментации сети
- отсутствию мониторинга безопасности
- отсутствию защиты почты
- отсутствию контроля привилегированных учетных записей

ДАННЫЕ АУДИТА:

{safe_results}

СФЕРА ДЕЯТЕЛЬНОСТИ:

{c_info.get("Сфера деятельности", "")}

КАТАЛОГ РЕШЕНИЙ КОМПАНИИ:

{vendor_context}

РЕГУЛЯТОРНЫЕ ТРЕБОВАНИЯ:

{regulator_context}

Верни ТОЛЬКО JSON.

Формат:

[
  {{
    "level": "HIGH",
    "risk": "Название риска",
    "description": "Описание риска",
    "impact": "Последствия для бизнеса",
    "recommendation": "Что необходимо сделать",
    "vendors": [
      "Vendor1",
      "Vendor2",
      "Vendor3"
    ],
    "regulators": [
      "ISO 27001"
    ]
  }}
]
"""

        response = model.generate_content(
            prompt,
            generation_config={
                "response_mime_type": "application/json"
            }
        )

        return json.loads(response.text)

    except Exception as e:
        st.error(f"Ошибка ИИ: {e}")
        return []



# --- AI BLOCK END ---

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

def inject_audit_design():
    st.markdown("""
    <style>
    :root {
        --audit-bg: #f6f7f9;
        --audit-panel: #ffffff;
        --audit-border: #d8dee8;
        --audit-text: #151922;
        --audit-muted: #667085;
        --audit-accent: #0f766e;
        --audit-accent-soft: #dff3ef;
        --audit-warn: #b45309;
        --audit-warn-soft: #fff7ed;
        --audit-risk: #b42318;
        --audit-risk-soft: #fff1f0;
        --audit-ink: #1f2937;
    }

    .stApp {
        background: var(--audit-bg);
        color: var(--audit-text);
    }

    section[data-testid="stSidebar"] {
        background: #fbfcfe;
        border-right: 1px solid var(--audit-border);
    }

    div[data-testid="stHeader"] {
        background: rgba(246, 247, 249, 0.92);
        backdrop-filter: blur(10px);
    }

    .block-container {
        padding-top: 2.2rem;
        padding-bottom: 7rem;
        max-width: 1380px;
    }

    h1, h2, h3 {
        letter-spacing: 0;
    }

    .audit-hero {
        background: var(--audit-panel);
        border: 1px solid var(--audit-border);
        border-radius: 8px;
        padding: 22px 24px;
        margin: 0 0 14px 0;
        box-shadow: 0 10px 30px rgba(16, 24, 40, 0.05);
    }

    .audit-brand-hero {
        display: grid;
        grid-template-columns: minmax(0, 1fr) minmax(300px, 380px);
        gap: 26px;
        align-items: center;
    }

    .audit-kicker {
        color: var(--audit-accent);
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0;
        margin-bottom: 8px;
    }

    .audit-title {
        color: var(--audit-text);
        font-size: 34px;
        font-weight: 760;
        line-height: 1.12;
        margin: 0 0 8px 0;
    }

    .audit-subtitle {
        color: var(--audit-muted);
        font-size: 15px;
        line-height: 1.55;
        max-width: 900px;
        margin: 0;
    }

    .audit-start-grid {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 10px;
        margin-top: 18px;
        max-width: 920px;
    }

    .audit-start-card {
        background: #f8fafc;
        border: 1px solid var(--audit-border);
        border-radius: 8px;
        padding: 12px 13px;
        min-height: 86px;
    }

    .audit-start-card strong {
        display: block;
        color: var(--audit-text);
        font-size: 13px;
        margin-bottom: 5px;
    }

    .audit-start-card span {
        color: var(--audit-muted);
        display: block;
        font-size: 12px;
        line-height: 1.45;
    }

    .audit-logo-lockup {
        border-left: 1px solid var(--audit-border);
        padding-left: 24px;
        text-align: center;
    }

    .audit-logo-lockup img {
        max-width: 320px;
        max-height: 120px;
        width: auto;
        height: auto;
        object-fit: contain;
    }

    .audit-logo-lockup .brand-name {
        color: var(--audit-text);
        font-size: 13px;
        font-weight: 760;
        margin-top: 10px;
    }

    .audit-logo-lockup .brand-signature {
        color: var(--audit-accent);
        font-size: 12px;
        font-weight: 760;
        margin-top: 3px;
    }

    .audit-section {
        background: var(--audit-panel);
        border: 1px solid var(--audit-border);
        border-radius: 8px;
        padding: 16px 18px;
        margin: 24px 0 14px 0;
    }

    .audit-section .eyebrow {
        color: var(--audit-muted);
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        margin-bottom: 4px;
    }

    .audit-section .title {
        color: var(--audit-text);
        font-size: 22px;
        font-weight: 760;
        margin-bottom: 4px;
    }

    .audit-section .body {
        color: var(--audit-muted);
        font-size: 14px;
        margin: 0;
    }

    .metric-card {
        background: var(--audit-panel);
        border: 1px solid var(--audit-border);
        border-radius: 8px;
        padding: 14px 14px 12px 14px;
        min-height: 96px;
    }

    .metric-card .label {
        color: var(--audit-muted);
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        margin-bottom: 8px;
    }

    .metric-card .value {
        color: var(--audit-text);
        font-size: 26px;
        font-weight: 760;
        line-height: 1.05;
    }

    .metric-card .hint {
        color: var(--audit-muted);
        font-size: 12px;
        margin-top: 6px;
    }

    .domain-row {
        display: grid;
        grid-template-columns: repeat(5, minmax(0, 1fr));
        gap: 10px;
        margin: 12px 0 18px 0;
    }

    .domain-card {
        background: var(--audit-panel);
        border: 1px solid var(--audit-border);
        border-radius: 8px;
        padding: 12px;
        min-height: 86px;
    }

    .domain-card strong {
        display: block;
        color: var(--audit-text);
        font-size: 13px;
        margin-bottom: 10px;
    }

    .domain-score {
        color: var(--audit-accent);
        font-size: 24px;
        font-weight: 760;
    }

    .risk-chip {
        border-radius: 8px;
        padding: 10px 12px;
        border: 1px solid var(--audit-border);
        background: var(--audit-panel);
        color: var(--audit-ink);
        margin-bottom: 8px;
        font-size: 13px;
        line-height: 1.45;
    }

    .risk-chip.critical {
        background: var(--audit-risk-soft);
        border-color: #fecdca;
    }

    .risk-chip.warn {
        background: var(--audit-warn-soft);
        border-color: #fed7aa;
    }

    .analysis-teaser {
        background: var(--audit-panel);
        border: 1px solid var(--audit-border);
        border-radius: 8px;
        padding: 16px 18px;
        margin: 20px 0 12px 0;
        box-shadow: 0 10px 26px rgba(16, 24, 40, 0.04);
    }

    .analysis-teaser-head {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 18px;
        margin-bottom: 14px;
    }

    .analysis-teaser-title {
        color: var(--audit-text);
        font-size: 19px;
        font-weight: 760;
        line-height: 1.25;
        margin-bottom: 4px;
    }

    .analysis-teaser-copy {
        color: var(--audit-muted);
        font-size: 13px;
        line-height: 1.45;
        max-width: 760px;
    }

    .analysis-pill {
        background: var(--audit-accent-soft);
        color: #0f5f59;
        border: 1px solid #b9e5de;
        border-radius: 999px;
        padding: 7px 10px;
        font-size: 12px;
        font-weight: 760;
        white-space: nowrap;
    }

    .teaser-grid {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 10px;
        margin-bottom: 14px;
    }

    .teaser-card {
        background: #f8fafc;
        border: 1px solid var(--audit-border);
        border-radius: 8px;
        padding: 12px;
        min-height: 86px;
    }

    .teaser-card .label {
        color: var(--audit-muted);
        font-size: 12px;
        font-weight: 700;
        margin-bottom: 7px;
    }

    .teaser-card .value {
        color: var(--audit-text);
        font-size: 23px;
        font-weight: 760;
        line-height: 1.05;
    }

    .teaser-card .hint {
        color: var(--audit-muted);
        font-size: 12px;
        margin-top: 6px;
    }

    .teaser-columns {
        display: grid;
        grid-template-columns: 1.1fr 0.9fr;
        gap: 12px;
    }

    .teaser-columns .label {
        color: var(--audit-muted);
        font-size: 12px;
        font-weight: 700;
        margin-bottom: 8px;
    }

    .mini-signal {
        background: #ffffff;
        border: 1px solid var(--audit-border);
        border-radius: 8px;
        padding: 10px 12px;
        margin-bottom: 8px;
        color: var(--audit-ink);
        font-size: 13px;
        line-height: 1.45;
    }

    .mini-signal.critical {
        border-color: #fecdca;
        background: var(--audit-risk-soft);
    }

    .mini-signal.warn {
        border-color: #fed7aa;
        background: var(--audit-warn-soft);
    }

    .mini-domain {
        display: flex;
        justify-content: space-between;
        gap: 10px;
        border-bottom: 1px solid #eef2f6;
        padding: 9px 0;
        color: var(--audit-ink);
        font-size: 13px;
    }

    .mini-domain:last-child {
        border-bottom: 0;
    }

    .mini-domain strong {
        color: var(--audit-text);
    }

    .analysis-teaser-note {
        color: var(--audit-muted);
        font-size: 12px;
        margin-top: 10px;
    }

    .analysis-status-panel {
        background: var(--audit-panel);
        border: 1px solid var(--audit-border);
        border-left: 4px solid var(--audit-accent);
        border-radius: 8px;
        padding: 16px 18px;
        margin-bottom: 16px;
        color: var(--audit-ink);
    }

    .analysis-status-title {
        color: var(--audit-text);
        font-size: 17px;
        font-weight: 760;
        margin-bottom: 6px;
    }

    .page-lock-note {
        color: var(--audit-risk);
        font-size: 13px;
        font-weight: 700;
        margin-top: 8px;
    }

    .analysis-log {
        background: #f8fafc;
        color: #344054;
        font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
        padding: 13px 14px;
        border: 1px solid var(--audit-border);
        min-height: 110px;
        border-radius: 8px;
        margin-bottom: 16px;
        font-size: 13px;
        line-height: 1.55;
    }

    .facts-panel {
        background: #fff8ed;
        border: 1px solid #fed7aa;
        border-left: 4px solid var(--audit-warn);
        border-radius: 8px;
        color: #1d2939;
        padding: 16px 18px;
        margin: 12px 0 18px 0;
        line-height: 1.55;
    }

    .facts-panel strong {
        color: #111827;
    }

    .section-feedback {
        background: #fff7ed;
        border: 1px solid #fed7aa;
        border-left: 4px solid var(--audit-warn);
        border-radius: 8px;
        color: #1d2939;
        margin: 12px 0 16px 0;
        padding: 12px 14px;
    }

    .section-feedback.ok {
        background: #f0fdf9;
        border-color: #99f6e4;
        border-left-color: var(--audit-accent);
    }

    .section-feedback-title {
        color: var(--audit-text);
        font-size: 13px;
        font-weight: 760;
        margin-bottom: 6px;
    }

    .section-feedback ul {
        margin: 0;
        padding-left: 18px;
    }

    .section-feedback li {
        font-size: 13px;
        line-height: 1.45;
        margin: 3px 0;
    }

    .sidebar-step {
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 7px 0;
        color: #344054;
        font-size: 13px;
    }

    .sidebar-dot {
        width: 9px;
        height: 9px;
        border-radius: 50%;
        flex: 0 0 9px;
    }

    .sidebar-dot.green {
        background: var(--audit-accent);
    }

    .sidebar-dot.red {
        background: var(--audit-risk);
    }

    .sidebar-dot.gray {
        background: #98a2b3;
    }

    .st-key-floating_draft_save {
        position: fixed;
        left: 50%;
        bottom: 18px;
        transform: translateX(-50%);
        z-index: 9999;
        width: min(560px, calc(100vw - 32px));
        padding: 0;
        background: transparent;
        border: 0;
        box-shadow: none;
    }

    .st-key-floating_draft_save .element-container,
    .st-key-floating_draft_save iframe {
        display: block !important;
        width: 100% !important;
        min-width: 0 !important;
        height: 96px !important;
        min-height: 96px !important;
        margin: 0 !important;
    }

    @media (max-width: 900px) {
        .audit-title {
            font-size: 26px;
        }

        .audit-brand-hero {
            grid-template-columns: 1fr;
        }

        .audit-logo-lockup {
            border-left: 0;
            border-top: 1px solid var(--audit-border);
            padding-left: 0;
            padding-top: 14px;
            text-align: center;
        }

        .audit-logo-lockup img {
            max-width: min(100%, 330px);
            max-height: 126px;
        }

        .audit-start-grid,
        .teaser-grid,
        .teaser-columns {
            grid-template-columns: 1fr;
        }

        .analysis-teaser-head {
            display: block;
        }

        .analysis-pill {
            display: inline-block;
            margin-top: 10px;
        }

        .domain-row {
            grid-template-columns: 1fr;
        }

        .st-key-floating_draft_save {
            bottom: 12px;
            width: calc(100vw - 24px);
        }
    }
    </style>
    """, unsafe_allow_html=True)


def get_logo_data_uri(path="logo.png"):
    if not os.path.exists(path):
        return ""

    with open(path, "rb") as logo_file:
        encoded = base64.b64encode(logo_file.read()).decode("ascii")

    return f"data:image/png;base64,{encoded}"


def render_app_header():
    logo_uri = get_logo_data_uri()
    logo_html = (
        f'<img src="{logo_uri}" alt="Khalil Trade">'
        if logo_uri
        else '<strong>Khalil Trade</strong>'
    )

    st.markdown(f"""
    <div class="audit-hero audit-brand-hero">
        <div>
            <div class="audit-kicker">Технический ИТ- и ИБ-аудит</div>
            <div class="audit-title">Опросник технического аудита ИТ и ИБ</div>
            <p class="audit-subtitle">
                Структурированный сбор данных для экспертной оценки инфраструктуры,
                зрелости защиты и подготовки XLSX-отчета.
            </p>
            <div class="audit-start-grid">
                <div class="audit-start-card">
                    <strong>7-10 минут на заполнение</strong>
                    <span>Начните с компании и конечных точек, остальные блоки включайте по факту наличия.</span>
                </div>
                <div class="audit-start-card">
                    <strong>Живая оценка по ходу анкеты</strong>
                    <span>Навигатор показывает готовность, слабые места и разделы, которые требуют внимания.</span>
                </div>
                <div class="audit-start-card">
                    <strong>XLSX-отчет для обсуждения</strong>
                    <span>На выходе формируется экспертная сводка, риски, домены защиты и быстрые улучшения.</span>
                </div>
            </div>
        </div>
        <div class="audit-logo-lockup">
            {logo_html}
            <div class="brand-name">Khalil Audit System v10.5</div>
            <div class="brand-signature">by Ivan Rudoy</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_section_marker(kicker, title, body):
    st.markdown(f"""
    <div class="audit-section">
        <div class="eyebrow">{html.escape(kicker)}</div>
        <div class="title">{html.escape(title)}</div>
        <p class="body">{html.escape(body)}</p>
    </div>
    """, unsafe_allow_html=True)


def get_section_errors(validation_errors, *markers):
    selected = []
    marker_values = [str(marker).lower() for marker in markers]

    for error in validation_errors:
        error_text = str(error)
        normalized = error_text.lower()
        if any(marker in normalized for marker in marker_values):
            selected.append(error_text)

    return list(dict.fromkeys(selected))


def render_section_feedback(title, errors, enabled=True):
    if not enabled:
        return

    if errors:
        items = "".join(
            f"<li>{html.escape(error)}</li>"
            for error in errors
        )
        st.markdown(f"""
        <div class="section-feedback">
            <div class="section-feedback-title">{html.escape(title)}: что исправить</div>
            <ul>{items}</ul>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="section-feedback ok">
            <div class="section-feedback-title">{html.escape(title)}: раздел выглядит корректно</div>
            <ul><li>Можно переходить к следующему блоку или уточнить детали в примечании.</li></ul>
        </div>
        """, unsafe_allow_html=True)


def draft_safe_value(value):
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value

    if isinstance(value, tuple):
        return [draft_safe_value(item) for item in value]

    if isinstance(value, list):
        return [draft_safe_value(item) for item in value]

    if isinstance(value, dict):
        return {
            str(key): draft_safe_value(item)
            for key, item in value.items()
            if draft_safe_value(item) is not None
        }

    return None


def collect_draft_state():
    excluded_keys = {
        "draft_upload",
        "cached_report_bytes",
        "report_ready",
        "generation_active",
        "draft_link_ready",
        "draft_link_notice",
        "_draft_query_marker",
    }
    draft_state = {}

    for key, value in st.session_state.items():
        if key in excluded_keys or str(key).startswith("FormSubmitter:"):
            continue

        safe_value = draft_safe_value(value)
        if safe_value is not None:
            draft_state[str(key)] = safe_value

    return {
        "schema": "khalil-audit-draft-v1",
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "state": draft_state,
    }


def build_draft_download():
    draft_payload = collect_draft_state()
    draft_bytes = json.dumps(
        draft_payload,
        ensure_ascii=False,
        indent=2,
    ).encode("utf-8")
    draft_filename = f"khalil_audit_draft_{datetime.now():%Y%m%d_%H%M}.json"

    return draft_bytes, draft_filename


def encode_draft_token(payload):
    raw = json.dumps(
        payload,
        ensure_ascii=False,
        separators=(",", ":"),
    ).encode("utf-8")
    compressed = zlib.compress(raw, level=9)
    return base64.urlsafe_b64encode(compressed).decode("ascii").rstrip("=")


def decode_draft_token(token):
    padding = "=" * (-len(token) % 4)
    compressed = base64.urlsafe_b64decode((token + padding).encode("ascii"))
    raw = zlib.decompress(compressed)
    return json.loads(raw.decode("utf-8"))


def apply_draft_state(payload):
    state = payload.get("state", payload)
    if not isinstance(state, dict):
        raise ValueError("Файл черновика не содержит блок state.")

    skipped_keys = {"draft_upload", "cached_report_bytes", "report_ready", "generation_active"}
    applied = 0

    for key, value in state.items():
        if key in skipped_keys or str(key).startswith("FormSubmitter:"):
            continue
        if key == "client_phone_code" and isinstance(value, list):
            value = tuple(value)
        st.session_state[key] = value
        applied += 1

    return applied


def restore_draft_from_query():
    token = st.query_params.get("draft")
    if not token:
        return

    if isinstance(token, list):
        token = token[0] if token else ""

    if not token:
        return

    marker = f"draft:{token[:24]}:{len(token)}"
    if st.session_state.get("_draft_query_marker") == marker:
        return

    try:
        payload = decode_draft_token(token)
        applied = apply_draft_state(payload)
        st.session_state["_draft_query_marker"] = marker
        st.session_state["draft_notice"] = (
            f"Черновик из ссылки применен: восстановлено полей {applied}."
        )
        st.query_params.clear()
        st.rerun()
    except Exception as exc:
        st.error(f"Не удалось применить черновик из ссылки: {exc}")


def render_floating_draft_save():
    draft_bytes, draft_filename = build_draft_download()
    draft_token = encode_draft_token(collect_draft_state())
    draft_text_json = json.dumps(draft_bytes.decode("utf-8"))
    draft_filename_json = json.dumps(draft_filename)
    draft_token_json = json.dumps(draft_token)
    can_share_json = "true" if len(draft_token) <= 6000 else "false"

    with st.container(key="floating_draft_save"):
        components.html(f"""
        <style>
        * {{
            box-sizing: border-box;
        }}

        .floating-draft-panel {{
            width: 100%;
            padding: 12px 14px 10px 14px;
            background: rgba(255, 255, 255, 0.97);
            border: 1px solid rgba(208, 213, 221, 0.92);
            border-radius: 22px;
            box-shadow: 0 18px 44px rgba(16, 24, 40, 0.16);
            backdrop-filter: blur(10px);
            font-family: "Inter", "Segoe UI", sans-serif;
        }}

        .floating-draft-actions {{
            display: grid;
            grid-template-columns: 1.35fr 1fr;
            gap: 12px;
            align-items: center;
        }}

        .floating-draft-action {{
            width: 100%;
            height: 52px;
            border-radius: 999px;
            border: 1px solid transparent;
            color: #ffffff;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 16px;
            font-weight: 760;
            line-height: 1;
            margin: 0;
            padding: 0 18px;
            white-space: nowrap;
        }}

        .floating-draft-action.save {{
            background: #f97316;
            border-color: #ea580c;
            box-shadow: 0 12px 28px rgba(234, 88, 12, 0.22);
        }}

        .floating-draft-action.save:hover {{
            background: #ea580c;
            border-color: #c2410c;
        }}

        .floating-draft-action.share {{
            background: #0f766e;
            border-color: #0f766e;
            box-shadow: 0 12px 28px rgba(15, 118, 110, 0.18);
        }}

        .floating-draft-action.share:hover {{
            background: #0d5f59;
            border-color: #0d5f59;
        }}

        .floating-draft-action:disabled {{
            background: #98a2b3;
            border-color: #98a2b3;
            cursor: not-allowed;
            box-shadow: none;
        }}

        .floating-draft-hint {{
            color: #475467;
            font-size: 11px;
            line-height: 1.25;
            margin-top: 8px;
            text-align: center;
            text-shadow: 0 1px 2px rgba(255, 255, 255, 0.85);
        }}

        @media (max-width: 460px) {{
            .floating-draft-panel {{
                padding: 10px 10px 8px 10px;
                border-radius: 18px;
            }}

            .floating-draft-actions {{
                gap: 8px;
            }}

            .floating-draft-action {{
                height: 46px;
                font-size: 13px;
                padding: 0 10px;
            }}
        }}
        </style>
        <div class="floating-draft-panel">
            <div class="floating-draft-actions">
                <button class="floating-draft-action save" id="floating-save" type="button">
                    Сохранить черновик
                </button>
                <button class="floating-draft-action share" id="floating-share" type="button">
                    Поделиться
                </button>
            </div>
            <div class="floating-draft-hint" id="floating-draft-status">
                Сохраните JSON-файл или поделитесь ссылкой на заполненную анкету.
            </div>
        </div>
        <script>
        const draftText = {draft_text_json};
        const draftFilename = {draft_filename_json};
        const draftToken = {draft_token_json};
        const canShareDraft = {can_share_json};
        const saveButton = document.getElementById("floating-save");
        const shareButton = document.getElementById("floating-share");
        const status = document.getElementById("floating-draft-status");

        function getParentUrl() {{
            try {{
                return window.parent.location.href;
            }} catch (error) {{
                return document.referrer || window.location.href;
            }}
        }}

        function buildShareUrl() {{
            const url = new URL(getParentUrl());
            url.searchParams.set("draft", draftToken);
            return url.toString();
        }}

        function saveDraftFile() {{
            const blob = new Blob([draftText], {{
                type: "application/json;charset=utf-8"
            }});
            const href = URL.createObjectURL(blob);
            const link = document.createElement("a");
            link.href = href;
            link.download = draftFilename;
            document.body.appendChild(link);
            link.click();
            link.remove();
            window.setTimeout(() => URL.revokeObjectURL(href), 1000);
            status.textContent = "JSON-черновик скачан. Его можно загрузить в блоке «Черновик анкеты».";
        }}

        async function copyShareUrl(shareUrl) {{
            if (navigator.clipboard && window.isSecureContext) {{
                await navigator.clipboard.writeText(shareUrl);
                status.textContent = "Ссылка скопирована. Передайте ее коллеге.";
                return;
            }}

            window.prompt("Скопируйте ссылку на заполненную анкету", shareUrl);
            status.textContent = "Скопируйте ссылку из окна браузера.";
        }}

        async function shareDraft() {{
            if (!canShareDraft) {{
                status.textContent = "Анкета слишком большая для ссылки. Скачайте JSON-черновик.";
                return;
            }}

            const shareUrl = buildShareUrl();
            const shareData = {{
                title: "Khalil Audit System",
                text: "Заполненная анкета Khalil Audit System by Ivan Rudoy",
                url: shareUrl
            }};

            try {{
                if (navigator.share) {{
                    await navigator.share(shareData);
                    status.textContent = "Анкета готова к отправке.";
                    return;
                }}

                await copyShareUrl(shareUrl);
            }} catch (error) {{
                if (error && error.name === "AbortError") {{
                    status.textContent = "Отправка отменена.";
                    return;
                }}

                try {{
                    await copyShareUrl(shareUrl);
                }} catch (copyError) {{
                    window.prompt("Скопируйте ссылку на заполненную анкету", shareUrl);
                    status.textContent = "Скопируйте ссылку из окна браузера.";
                }}
            }}
        }}

        if (!canShareDraft) {{
            shareButton.disabled = true;
            shareButton.title = "Анкета слишком большая для ссылки";
        }}

        saveButton.addEventListener("click", saveDraftFile);
        shareButton.addEventListener("click", shareDraft);
        </script>
        """, height=96)


def render_draft_share_button(
    draft_token,
    label="Поделиться заполненной анкетой",
    status_text="Откроет меню отправки или скопирует ссылку.",
    height=82,
    compact=False,
):
    token_json = json.dumps(draft_token)
    label_json = json.dumps(label)
    status_json = json.dumps(status_text)
    status_markup = (
        '<div class="draft-share-status" id="draft-share-status"></div>'
        if status_text
        else '<div class="draft-share-status sr-only" id="draft-share-status"></div>'
    )
    border_radius = "999px" if compact else "8px"
    min_height = "42px" if compact else "38px"
    font_size = "14px" if compact else "14px"

    components.html(f"""
    <style>
    .draft-share-wrap {{
        font-family: "Inter", "Segoe UI", sans-serif;
        width: 100%;
    }}

    .draft-share-button {{
        width: 100%;
        min-height: {min_height};
        border: 1px solid #0f766e;
        border-radius: {border_radius};
        background: #0f766e;
        color: #ffffff;
        cursor: pointer;
        font-size: {font_size};
        font-weight: 650;
        line-height: 1.2;
    }}

    .draft-share-button:hover {{
        background: #0d5f59;
        border-color: #0d5f59;
    }}

    .draft-share-status {{
        color: #667085;
        font-size: 11px;
        line-height: 1.35;
        margin-top: 7px;
        text-align: center;
    }}

    .sr-only {{
        position: absolute;
        width: 1px;
        height: 1px;
        padding: 0;
        margin: -1px;
        overflow: hidden;
        clip: rect(0, 0, 0, 0);
        white-space: nowrap;
        border: 0;
    }}
    </style>
    <div class="draft-share-wrap">
        <button class="draft-share-button" id="draft-share-button" type="button">
        </button>
        {status_markup}
    </div>
    <script>
    const draftToken = {token_json};
    const buttonLabel = {label_json};
    const initialStatus = {status_json};
    const button = document.getElementById("draft-share-button");
    const status = document.getElementById("draft-share-status");
    button.textContent = buttonLabel;
    if (status) {{
        status.textContent = initialStatus;
    }}

    function getParentUrl() {{
        try {{
            return window.parent.location.href;
        }} catch (error) {{
            return document.referrer || window.location.href;
        }}
    }}

    function buildShareUrl() {{
        const url = new URL(getParentUrl());
        url.searchParams.set("draft", draftToken);
        return url.toString();
    }}

    async function copyShareUrl(shareUrl) {{
        if (navigator.clipboard && window.isSecureContext) {{
            await navigator.clipboard.writeText(shareUrl);
            status.textContent = "Ссылка скопирована. Передайте ее коллеге.";
            return;
        }}

        window.prompt("Скопируйте ссылку на заполненную анкету", shareUrl);
        status.textContent = "Скопируйте ссылку из окна браузера.";
    }}

    button.addEventListener("click", async () => {{
        const shareUrl = buildShareUrl();
        const shareData = {{
            title: "Khalil Audit System",
            text: "Заполненная анкета Khalil Audit System by Ivan Rudoy",
            url: shareUrl
        }};

        try {{
            if (navigator.share) {{
                await navigator.share(shareData);
                status.textContent = "Анкета готова к отправке.";
                return;
            }}

            await copyShareUrl(shareUrl);
        }} catch (error) {{
            if (error && error.name === "AbortError") {{
                status.textContent = "Отправка отменена.";
                return;
            }}

            try {{
                await copyShareUrl(shareUrl);
            }} catch (copyError) {{
                window.prompt("Скопируйте ссылку на заполненную анкету", shareUrl);
                status.textContent = "Скопируйте ссылку из окна браузера.";
            }}
        }}
    }});
    </script>
    """, height=height)


def render_draft_tools():
    if st.session_state.get("draft_notice"):
        st.success(st.session_state.pop("draft_notice"))

    with st.expander("Черновик анкеты: сохранить или продолжить заполнение", expanded=False):
        st.caption(
            "Сохраните JSON-черновик или создайте ссылку, если анкету нужно передать коллеге "
            "и продолжить заполнение на другом компьютере. Черновик содержит введенные ответы, "
            "поэтому передавайте его только доверенному получателю."
        )

        draft_bytes, draft_filename = build_draft_download()
        draft_token = encode_draft_token(collect_draft_state())

        col_save, col_link, col_load = st.columns(3)
        with col_save:
            st.download_button(
                "Скачать черновик JSON",
                data=draft_bytes,
                file_name=draft_filename,
                mime="application/json",
                use_container_width=True,
                key="draft_expander_download",
            )

        with col_link:
            if len(draft_token) > 6000:
                st.warning(
                    "Анкета уже слишком большая для надежной ссылки. "
                    "Скачайте JSON-черновик и передайте файл коллеге."
                )
            else:
                render_draft_share_button(draft_token)

        with col_load:
            uploaded_draft = st.file_uploader(
                "Загрузить черновик JSON",
                type=["json"],
                key="draft_upload",
                help="После применения страница перезагрузится и подставит сохраненные ответы.",
            )
            if uploaded_draft and st.button("Применить черновик", use_container_width=True):
                try:
                    payload = json.loads(uploaded_draft.getvalue().decode("utf-8"))
                    applied = apply_draft_state(payload)
                    st.session_state["draft_notice"] = f"Черновик применен: восстановлено полей {applied}."
                    st.rerun()
                except Exception as exc:
                    st.error(f"Не удалось применить черновик: {exc}")


def render_generation_guard(active):
    action = "install" if active else "remove"
    components.html(f"""
    <script>
    const target = window.parent;
    const message = "Экспертный отчет формируется. Это может занять до 4 минут. Не закрывайте и не обновляйте страницу.";

    if (!target.__khalilBeforeUnloadHandler) {{
      target.__khalilBeforeUnloadHandler = (event) => {{
        event.preventDefault();
        event.returnValue = message;
        return message;
      }};
    }}

    if ("{action}" === "install" && !target.__khalilBeforeUnloadActive) {{
      target.addEventListener("beforeunload", target.__khalilBeforeUnloadHandler);
      target.__khalilBeforeUnloadActive = true;
    }}

    if ("{action}" === "remove" && target.__khalilBeforeUnloadActive) {{
      target.removeEventListener("beforeunload", target.__khalilBeforeUnloadHandler);
      target.__khalilBeforeUnloadActive = false;
    }}
    </script>
    """, height=0)


inject_audit_design()

# Якорь для принудительного перехода в начало страницы
st.markdown("<div id='top'></div>", unsafe_allow_html=True)

# --- НАСТРОЙКИ TELEGRAM (из Secrets) ---
TOKEN = get_app_secret("TELEGRAM_TOKEN")
CHAT_ID = get_app_secret("TELEGRAM_CHAT_ID")

restore_draft_from_query()

render_app_header()
render_floating_draft_save()

# --- ИНСТРУКЦИЯ ДЛЯ ПОЛЬЗОВАТЕЛЯ ---
with st.expander("Инструкция по заполнению", expanded=False):
    st.markdown("""
    ### Как пройти анкету

    1. **Начните с обязательного минимума:** заполните общую информацию о компании и блок **Конечные точки (АРМ)**. У заказчика почти всегда есть рабочие станции, поэтому этот блок не отключается и нужен для получения отчета.
    2. **Включайте только реальные разделы:** сеть, серверы, виртуализация, СХД, внутренние системы, ИБ и разработка открываются тумблерами. Если раздел включен, его вложенные обязательные поля нужно заполнить.
    3. **Смотрите навигатор слева:** красная точка означает, что раздел требует внимания, зеленая - раздел выглядит заполненным, серая - раздел отключен и не участвует в анкете.
    4. **Исправляйте по подсказкам на месте:** под каждым крупным блоком есть панель “что исправить”. Она показывает конкретные недостающие поля именно для этого раздела.
    5. **Используйте предварительную аналитику:** блок “Предварительная аналитика” показывает сводку аудита и быстрые улучшения еще до формирования финального XLSX. Полную версию можно раскрыть ниже.
    6. **Сохраняйте черновик при совместном заполнении:** скачайте JSON-черновик и передайте его коллеге. Он сможет загрузить файл в этом же блоке и продолжить заполнение с сохраненных ответов.
    7. **Не закрывайте страницу во время отчета:** формирование экспертного отчета может занять до 4 минут. Пока идет генерация, не обновляйте и не закрывайте вкладку.

    Поля “Примечание” необязательны, но помогают эксперту точнее описать контекст, ограничения и планы развития инфраструктуры.
    """)

render_draft_tools()

data = {}
client_info = {}
validation_errors = []
score = 0
# --- ШАПКА: ОБЩАЯ ИНФОРМАЦИЯ ---
render_section_marker(
    "01 / КОМПАНИЯ",
    "Общая информация",
    "Контекст компании, отрасль и контактные данные для экспертного отчета."
)
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город*", key="client_city", help="Укажите город фактического нахождения головного офиса.")
    industry_options = ["Финтех / Банки", "Ритейл / E-commerce", "Производство", "IT / Разработка", "Госсектор", "Другое"]
    selected_ind = st.selectbox(
        "Сфера деятельности компании*",
        [""] + industry_options,
        format_func=lambda x: "Выберите сферу..." if x == "" else x,
        key="client_industry_select",
        help="Отрасль влияет на профиль угроз и регуляторные требования."
    )

    if selected_ind == "Другое":
        industry = st.text_input("Укажите вашу сферу деятельности*", key="client_industry_other", help="Введите отрасль вручную")
    else:
        industry = selected_ind
    
    client_info['Сфера деятельности'] = industry
    client_info['Наименование компании'] = st.text_input("Наименование компании*", key="client_company_name", help="Официальное или сокращенное название юрлица.")

    site_input = st.text_input("Сайт компании*", key="site_field", placeholder="example.kz", help="Используется для анализа внешнего цифрового отпечатка.")
    client_info['Сайт компании'] = site_input

    custom_email_mode = st.checkbox("Email отличается от сайта", key="client_custom_email_mode", help="Отметьте, если корпоративная почта находится на другом домене.")

    if custom_email_mode:
        client_info['Email'] = st.text_input("Email контактного лица*", key="client_email_custom", help="Личный корпоративный email для отправки результатов.")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица (логин)*")
            e_col1, e_col2 = st.columns([2, 3])
            with e_col1:
                email_prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="email_pre", help="Только часть адреса до символа @")
            with e_col2:
                st.markdown(f"<div style='padding-top: 5px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
            client_info['Email'] = f"{email_prefix}@{clean_domain}" if email_prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица*", key="client_contact_name", help="С кем наш эксперт сможет обсудить детали отчета.")
    client_info['Должность'] = st.text_input("Должность*", key="client_contact_role", help="Например: ИТ-Директор, Системный администратор, CEO.")
    
    st.write("Контактный телефон*")
    p_col1, p_col2 = st.columns([1, 2])
    country_codes = [
        ("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"), ("🇰🇬 +996", "+996"),
        ("🇹🇯 +992", "+992"), ("🇦🇪 +971", "+971"), ("🇹🇷 +90", "+90"), ("🇦🇿 +994", "+994")
    ]
    selected_code = p_col1.selectbox("Код", country_codes, format_func=lambda x: x[0], label_visibility="collapsed", key="client_phone_code")
    phone_num = p_col2.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed", key="client_phone_number", help="Телефон для оперативной связи.")
    client_info['Контактный телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

if not all([client_info.get('Город'), client_info.get('Наименование компании'), client_info.get('Сфера деятельности'), client_info.get('Сайт компании'), client_info.get('Email'), client_info.get('ФИО контактного лица'), client_info.get('Должность'), phone_num]):
    validation_errors.append("Заполните все обязательные поля в блоке 'Общая информация'")

render_section_feedback(
    "Общая информация",
    get_section_errors(validation_errors, "общая информация")
)

st.divider()

# --- БЛОК 1: ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ ---
render_section_marker(
    "02 / ИНФРАСТРУКТУРА",
    "Информационные технологии",
    "АРМ, сеть, серверы, хранение данных и внутренние бизнес-системы."
)

st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт)", min_value=0, step=1, key="total_arm", help="Общее число ПК, ноутбуков и тонких клиентов в организации.")
data['1.1. Всего АРМ'] = total_arm
selected_os_arm = st.multiselect("ОС на АРМ", ["Windows XP/Vista/7/8", "Windows 10", "Windows 11", "Linux", "macOS", "Другое"], key="selected_os_arm", help="Выберите все типы операционных систем, используемых сотрудниками.")

sum_os_arm = 0
if selected_os_arm:
    for os_item in selected_os_arm:
        val = st.number_input(f"Кол-во на {os_item}", min_value=0, step=1, key=f"arm_{os_item}", help=f"Укажите точное или примерное число устройств с {os_item}.")
        data[f"ОС АРМ ({os_item})"] = val
        sum_os_arm += val

data['1.1. Примечание'] = st.text_area("Примечание к разделу 1.1", placeholder="Напр.: планируем обновление Windows 10 до 11 в Q3", key="note_1_1")

if total_arm <= 0:
    validation_errors.append("Укажите количество АРМ")
elif not selected_os_arm:
    validation_errors.append("Выберите ОС на АРМ")
elif sum_os_arm != total_arm:
    st.warning(f"⚠️ Ошибка: Сумма по ОС ({sum_os_arm}) должна быть равна общему количеству АРМ ({total_arm}).")
    validation_errors.append("Несовпадение количества АРМ и ОС")

render_section_feedback(
    "Конечные точки",
    get_section_errors(validation_errors, "арм", "ос на арм")
)

st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
main_speed, back_speed, ap_cnt = 0, 0, 0
selected_routing = []
ngfw_vendor = "Нет"
wifi_enabled = False
wifi_ctrl_enabled = False
net_active = st.toggle("Своя сетевая инфраструктура", key="net_toggle", help="Активируйте, если организация самостоятельно управляет сетевым оборудованием.")

if net_active:
    net_types = ["Оптика", "RJ45 (Ethernet)", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL/VDSL", "Нет"]
    routing_types = ["Статическая", "RIP", "OSPF", "EIGRP", "BGP", "IS-IS"]
    
    col_net1, col_net2 = st.columns(2)
    with col_net1:
        st.write("Основной канал")
        main_type = st.selectbox("Тип (основной)", net_types, key="main_net_type", index=7, help="Технология подключения основного интернет-канала.")
        main_speed = st.number_input("Скорость основного (Mbit/s)", min_value=0, step=10, key="main_net_speed")
        data['1.2.1. Основной канал'] = f"{main_type} ({main_speed} Mbit/s)"
    with col_net2:
        st.write("Резервный канал")
        back_type = st.selectbox("Тип (резервный)", net_types, index=7, key="back_net_type", help="Наличие и тип независимого резервного канала.")
        back_speed = st.number_input("Скорость резервного (Mbit/s)", min_value=0, step=10, key="back_net_speed")
        data['1.2.2. Резервный канал'] = f"{back_type} ({back_speed} Mbit/s)"

    st.write("Логика сети")
    selected_routing = st.multiselect("Тип маршрутизации*", routing_types, key="routing_sel", help="Протоколы динамической маршрутизации, используемые в сети.")
    data['1.2.3. Маршрутизация'] = ", ".join(selected_routing)
    if not selected_routing:
        validation_errors.append("Выберите тип маршрутизации")

    st.write("Активное сетевое оборудование")
    c_net1, c_net2, c_net3 = st.columns(3)
    with c_net1:
        if st.checkbox("Маршрутизаторы", key="router_chk", help="Устройства для связи разных сетей и выхода в интернет."):
            r_count = st.number_input("Кол-во маршрутизаторов", min_value=0, step=1, key="router_cnt")
            data['1.2.4. Маршрутизаторы'] = f"Да ({r_count} шт)"
            if r_count == 0: validation_errors.append("Укажите количество маршрутизаторов")
    with c_net2:
        if st.checkbox("Коммутаторы L2", key="swl2_chk", help="Управляемые или неуправляемые коммутаторы уровня доступа."):
            sw2_count = st.number_input("Кол-во коммутаторов L2", min_value=0, step=1, key="swl2_cnt")
            data['1.2.5. Коммутаторы L2'] = f"Да ({sw2_count} шт)"
            if sw2_count == 0: validation_errors.append("Укажите количество коммутаторов L2")
    with c_net3:
        if st.checkbox("Коммутаторы L3", key="swl3_chk", help="Коммутаторы с функциями маршрутизации (ядро или агрегация)."):
            sw3_count = st.number_input("Кол-во коммутаторов L3", min_value=0, step=1, key="swl3_cnt")
            data['1.2.6. Коммутаторы L3'] = f"Да ({sw3_count} шт)"
            if sw3_count == 0: validation_errors.append("Укажите количество коммутаторов L3")

    st.write("Уровни сети")
    l_col1, l_col2, l_col3 = st.columns(3)
    with l_col1:
        if st.checkbox("Ядро (Core)", key="net_core", help="Центральная часть сети, обеспечивающая максимальную скорость."):
            core_v = st.text_input("Основной производитель (Core)", key="core_vendor", help="Например: Cisco, Huawei, Juniper, MikroTik.")
            data['Уровень сети Ядро'] = core_v
            if not core_v: validation_errors.append("Укажите производителя Core-уровня")
    with l_col2:
        if st.checkbox("Уровень распределения", key="net_dist", help="Связующее звено между ядром и уровнем доступа."):
            dist_v = st.text_input("Основной производитель (Dist)", key="dist_vendor")
            data['Уровень сети Распределение'] = dist_v
            if not dist_v: validation_errors.append("Укажите производителя уровня распределения")
    with l_col3:
        if st.checkbox("Уровень доступа", key="net_acc", help="Уровень, к которому подключаются конечные пользователи."):
            acc_v = st.text_input("Основной производитель (Access)", key="acc_vendor")
            data['Уровень сети Доступ'] = acc_v
            if not acc_v: validation_errors.append("Укажите производителя уровня доступа")

    wifi_enabled = st.checkbox("Wi-Fi", key="wifi_toggle", help="Наличие корпоративной беспроводной сети.")
    if wifi_enabled:
        w_col1, w_col2, w_col3 = st.columns(3)
        with w_col1:
            wifi_ctrl_enabled = st.checkbox("Контроллер", key="wifi_ctrl", help="Централизованное управление точками доступа (аппаратное или программное).")
            if wifi_ctrl_enabled:
                wc_v = st.text_input("Производитель/модель контроллера", key="wc_vendor")
                data['Wi-Fi Контроллер'] = wc_v
                if not wc_v: validation_errors.append("Укажите модель Wi-Fi контроллера")
            else:
                data['Wi-Fi Контроллер'] = "Нет"
        with w_col2:
            ap_cnt = st.number_input("Количество точек доступа (шт)", min_value=0, step=1, key="ap_cnt", help="Общее число активных Wi-Fi точек.")
            data['Wi-Fi Точки доступа'] = ap_cnt
            if ap_cnt == 0: validation_errors.append("Укажите количество точек доступа Wi-Fi")
        with w_col3:
            wf_types = ["Wi-Fi 6/6E (802.11ax)", "Wi-Fi 5 (802.11ac)", "Wi-Fi 4 (802.11n)", "Другое"]
            data['Wi-Fi Тип'] = st.selectbox("Тип Wi-Fi", wf_types, key="wf_type_sel", help="Преимущественный стандарт беспроводной связи.")

    if st.checkbox("Межсетевой экран (NGFW)", key="ngfw_chk", help="Многофункциональные шлюзы безопасности (FortiGate, UserGate, CheckPoint и т.д.)."):
        ngfw_vendor = st.text_input("Производитель (NGFW)", key="ngfw_v")
        data['1.2.7. NGFW'] = f"Да ({ngfw_vendor if ngfw_vendor else 'не указан'})"
        if not ngfw_vendor: validation_errors.append("Укажите производителя NGFW")
        score += 20
    
    data['1.2. Примечание'] = st.text_area("Примечание к разделу 1.2", placeholder="Особенности топологии сети...", key="note_1_2")

render_section_feedback(
    "Сетевая инфраструктура",
    get_section_errors(
        validation_errors,
        "маршрутизац",
        "маршрутизатор",
        "коммутатор",
        "core",
        "распредел",
        "доступ",
        "wi-fi",
        "ngfw"
    ),
    enabled=net_active
)

st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
sum_os_srv = 0
phys_count = 0
virt_count = 0
server_active = st.toggle("Серверы и виртуализация", key="server_toggle", help="Включите, если у заказчика есть физические серверы, виртуальные машины или платформы виртуализации.")
v_n_b = "Нет"
selected_os_srv = []
selected_virt_sys = []

if server_active:
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        phys_count = st.number_input("Количество физических серверов", min_value=0, step=1, key="phys_srv", help="Количество 'железных' серверов в серверной или ЦОД.")
        data['1.3.1. Физические серверы'] = phys_count
    with col_s2:
        virt_count = st.number_input("Количество виртуальных серверов", min_value=0, step=1, key="virt_srv", help="Суммарное количество виртуальных машин (VM).")
        data['1.3.2. Виртуальные серверы'] = virt_count

    if phys_count == 0 and virt_count == 0:
        validation_errors.append("Укажите количество физических или виртуальных серверов")

    s_os_list = ["Windows Server 2008/2012 R2", "Windows Server 2016", "Windows Server 2019", "Windows Server 2022", "Linux", "Unix", "Другое"]
    selected_os_srv = st.multiselect("Выберите ОС серверов", s_os_list, key="ms_srv_list", help="Операционные системы, установленные на серверах.")
    if selected_os_srv:
        for os_s in selected_os_srv:
            val_os = st.number_input(f"Кол-во на {os_s}", min_value=0, key=f"fsrv_{os_s}")
            data[f"ОС Сервера ({os_s})"] = val_os
            sum_os_srv += val_os

    if virt_count > 0 and sum_os_srv < virt_count:
        st.warning(f"⚠️ Ошибка: Количество ОС ({sum_os_srv}) должно быть больше или равно количеству виртуальных серверов ({virt_count}).")
        validation_errors.append("Недостаточное количество ОС для серверов")

    selected_virt_sys = st.multiselect("Выберите системы виртуализации", ["VMware", "Hyper-V", "Proxmox", "KVM", "Другое", "Нет"], key="virt_sys_list", help="Технологии управления виртуальной инфраструктурой.")
    if selected_virt_sys and "Нет" not in selected_virt_sys:
        for v_sys in selected_virt_sys:
            v_h_cnt = st.number_input(f"Количество хостов {v_sys}", min_value=0, step=1, key=f"fv_cnt_{v_sys}", help=f"Сколько физических серверов (нод) в кластере {v_sys}?")
            data[f"Система виртуализации ({v_sys})"] = v_h_cnt
            if v_h_cnt == 0: validation_errors.append(f"Укажите количество хостов для {v_sys}")

    if st.checkbox("Резервное копирование", key="ib_backup", help="Наличие специализированного ПО для бэкапа (Veeam, Commvault, Veritas и т.д.)."):
        v_n_b = st.text_input("Вендор Резервного копирования", key="vn_backup", help="Укажите название используемого продукта.")
        data["Резервное копирование"] = v_n_b
        if not v_n_b: validation_errors.append("Укажите вендора резервного копирования")
        score += 20

    data['1.3. Примечание'] = st.text_area("Примечание к разделу 1.3", placeholder="Специфика серверного парка...", key="note_1_3")

render_section_feedback(
    "Серверы и виртуализация",
    get_section_errors(
        validation_errors,
        "сервер",
        "хост",
        "резервного копирования"
    ),
    enabled=server_active
)

st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
st_media_sel = []
cnt_hdd = 0
cnt_ssd = 0
raid_selected = []
storage_active = st.toggle("Есть собственная СХД", key="storage_toggle")
if storage_active:
    st_media_sel = st.multiselect(
        "Типы носителей",
        ["HDD (NL-SAS / SATA)", "SSD (SATA / SAS)", "NVMe", "SCM"],
        key="st_media"
    )
    data['1.4.1. Типы носителей'] = ", ".join(st_media_sel) if st_media_sel else "Не указано"

    col_pct1, col_pct2 = st.columns(2)
    with col_pct1:
        cnt_hdd = st.number_input("Количество дисков HDD", min_value=0, step=1, key="cnt_hdd")
        data['1.4.2. Кол-во HDD'] = cnt_hdd
    with col_pct2:
        cnt_ssd = st.number_input("Количество дисков SSD", min_value=0, step=1, key="cnt_ssd")
        data['1.4.3. Кол-во SSD'] = cnt_ssd

    if st_media_sel and (cnt_hdd + cnt_ssd == 0):
        st.info("ℹ️ Укажите количество дисков для СХД.")
        validation_errors.append("Не указано количество дисков СХД")

    col_chk1, col_chk2 = st.columns(2)
    with col_chk1:
        hybrid = st.checkbox("Используется гибридная СХД", key="hybrid_st")
        data['1.4.4. Гибридная СХД'] = "Да" if hybrid else "Нет"
    with col_chk2:
        allflash = st.checkbox("Есть All-Flash массивы", key="allflash_st")
        data['1.4.5. All-Flash'] = "Да" if allflash else "Нет"
        if allflash: score += 5

    raid_selected = st.multiselect(
        "Используемые RAID-группы",
        ["RAID 0", "RAID 1", "RAID 5", "RAID 6", "RAID 10", "RAID 50", "RAID 60", "JBOD"],
        key="raid_list"
    )
    data['1.4.6. RAID-группы'] = ", ".join(raid_selected) if raid_selected else "Не указано"

    if not raid_selected:
        validation_errors.append("Не указаны RAID-группы СХД")
    if "RAID 0" in raid_selected or "JBOD" in raid_selected:
        score -= 10

    data['1.4. Примечание'] = st.text_area("Примечание к разделу 1.4", placeholder="SAN/NAS, replication, snapshot, DR-site, tiering и т.д.", key="note_1_4")

render_section_feedback(
    "Системы хранения данных",
    get_section_errors(validation_errors, "схд", "raid"),
    enabled=storage_active
)

st.write("---")
st.subheader("1.5. Внутренние Информационные системы")
is_active = st.toggle("ИС организации", key="is_toggle", help="Бизнес-приложения и корпоративные сервисы.")
if is_active:
    is_types = {
        "ERP": "erp", "CRM": "crm", "HelpDesk/ServiceDesk": "sd", 
        "СЭД (Документооборот)": "sed", "HRM (Кадры)": "hrm", 
        "BI (Аналитика)": "bi", "WMS (Склад)": "wms", "Учет (Бухгалтерия)": "acc"
    }
    
    m_opts = ["Exchange (On-Prem)", "Lotus", "Microsoft 365", "Google Workspace", "Собственный", "Нет"]
    m_sys = st.selectbox("Почтовая система", m_opts, key="mail_system", help="Где физически и логически располагается ваша электронная почта.")
    
    if m_sys in ["Exchange (On-Prem)", "Lotus"]:
        m_ver = st.text_input(f"Версия {m_sys}*", key="mail_version_input", help="Например: 2016 CU23 или v9.0.1.")
        data['1.5.1. Почтовая система'] = f"{m_sys} (v.{m_ver})"
        if not m_ver: validation_errors.append(f"Укажите версию {m_sys}")
    else:
        data['1.5.1. Почтовая система'] = m_sys

    for label, ks in is_types.items():
        if st.checkbox(label, key=f"is_chk_{ks}"):
            c_is1, c_is2 = st.columns(2)
            with c_is1:
                name_is = st.text_input(f"Название продукта {label}*", key=f"name_{ks}")
            with c_is2:
                ver_is = st.text_input(f"Версия {label}*", key=f"ver_{ks}")
            data[f"ИС {label}"] = f"{name_is} (v.{ver_is})"
            if not name_is or not ver_is:
                validation_errors.append(f"Укажите название и версию для {label}")
    
    data['1.5. Примечание'] = st.text_area("Примечание к разделу 1.5", placeholder="Дополнительные ИС...", key="note_1_5")

render_section_feedback(
    "Внутренние информационные системы",
    get_section_errors(validation_errors, "версию", "название и версию"),
    enabled=is_active
)

st.divider()

# --- БЛОК 2: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ ---
render_section_marker(
    "03 / БЕЗОПАСНОСТЬ",
    "Информационная безопасность",
    "Текущие средства защиты, мониторинг, доступ, приложения и управление уязвимостями."
)

enable_security = st.toggle("Включить блок ИБ", value=False)

# Инициализация переменных ИБ
epp, epp_v, edr, edr_v, xdr, xdr_v, mdr, mdr_v = False, "", False, "", False, "", False, ""
dlp, dlp_v, mail_sec, mail_v, casb, casb_v = False, "", False, "", False, ""
waf, waf_v, ddos, ddos_v, ids, ids_v, nac, nac_v, ztna, ztna_v = False, "", False, "", False, "", False, "", False, ""
sast, sast_v, dast, dast_v = False, "", False, ""
iam, iam_v, mfa, mfa_v, pam, pam_v = False, "", False, "", False, ""
siem, siem_v, soar, soar_v = False, "", False, ""
vuln, vuln_v, patch, patch_v, nad, nad_v = False, "", False, "", False, ""

if enable_security:
    errors = []

    # =========================
    # Защита конечных устройств
    # =========================
    st.markdown("#### Защита конечных устройств")
    col1, col2 = st.columns(2)
    with col1:
        epp = st.checkbox("EPP (антивирусная защита)", key="epp")
        epp_v = st.text_input("Производитель EPP", key="epp_v") if epp else ""
        data['Блок 2. EPP'] = epp_v if epp else "Нет"
        
        edr = st.checkbox("EDR (обнаружение и реагирование)", key="edr")
        edr_v = st.text_input("Производитель EDR", key="edr_v") if edr else ""
        data['Блок 2. EDR'] = edr_v if edr else "Нет"
    with col2:
        xdr = st.checkbox("XDR (расширенная защита)", key="xdr")
        xdr_v = st.text_input("Производитель XDR", key="xdr_v") if xdr else ""
        data['Блок 2. XDR'] = xdr_v if xdr else "Нет"
        
        mdr = st.checkbox("MDR (внешний мониторинг)", key="mdr")
        mdr_v = st.text_input("Провайдер MDR", key="mdr_v") if mdr else ""
        data['Блок 2. MDR'] = mdr_v if mdr else "Нет"

    # =========================
    # Защита данных
    # =========================
    st.markdown("#### Защита данных")
    col1, col2 = st.columns(2)
    with col1:
        dlp = st.checkbox("DLP (предотвращение утечек)", key="dlp")
        dlp_v = st.text_input("Производитель DLP", key="dlp_v") if dlp else ""
        data['Блок 2. DLP'] = dlp_v if dlp else "Нет"
        
        mail_sec = st.checkbox("Mail Security (защита почты)", key="mail_sec")
        mail_v = st.text_input("Производитель Mail Security", key="mail_v") if mail_sec else ""
        data['Блок 2. Mail Security'] = mail_v if mail_sec else "Нет"
    with col2:
        casb = st.checkbox("CASB (контроль облаков)", key="casb")
        casb_v = st.text_input("Производитель CASB", key="casb_v") if casb else ""
        data['Блок 2. CASB'] = casb_v if casb else "Нет"

    # =========================
    # Сетевая безопасность
    # =========================
    st.markdown("#### Сетевая безопасность")
    col1, col2 = st.columns(2)
    with col1:
        waf = st.checkbox("WAF (защита веб-приложений)", key="waf")
        waf_v = st.text_input("Производитель WAF", key="waf_v") if waf else ""
        data['Блок 2. WAF'] = waf_v if waf else "Нет"
        
        ddos = st.checkbox("Anti-DDoS (защита от атак)", key="ddos")
        ddos_v = st.text_input("Производитель Anti-DDoS", key="ddos_v") if ddos else ""
        data['Блок 2. Anti-DDoS'] = ddos_v if ddos else "Нет"
        
        nad = st.checkbox("NAD (Network Attack Discovery)", key="nad")
        nad_v = st.text_input("Производитель NAD", key="nad_v") if nad else ""
        data['Блок 2. NAD'] = nad_v if nad else "Нет"
    with col2:
        ids = st.checkbox("IDS/IPS (сетевые атаки)", key="ids")
        ids_v = st.text_input("Производитель IDS/IPS", key="ids_v") if ids else ""
        data['Блок 2. IDS/IPS'] = ids_v if ids else "Нет"
        
        nac = st.checkbox("NAC (контроль доступа)", key="nac")
        nac_v = st.text_input("Производитель NAC", key="nac_v") if nac else ""
        data['Блок 2. NAC'] = nac_v if nac else "Нет"
        
        ztna = st.checkbox("ZTNA (Zero Trust доступ)", key="ztna")
        ztna_v = st.text_input("Производитель ZTNA", key="ztna_v") if ztna else ""
        data['Блок 2. ZTNA'] = ztna_v if ztna else "Нет"

    # =========================
    # Безопасность приложений
    # =========================
    st.markdown("#### Безопасность приложений")
    col1, col2 = st.columns(2)
    with col1:
        sast = st.checkbox("SAST (анализ кода)", key="sast")
        sast_v = st.text_input("Производитель SAST", key="sast_v") if sast else ""
        data['Блок 2. SAST'] = sast_v if sast else "Нет"
    with col2:
        dast = st.checkbox("DAST (тестирование приложений)", key="dast")
        dast_v = st.text_input("Производитель DAST", key="dast_v") if dast else ""
        data['Блок 2. DAST'] = dast_v if dast else "Нет"

    # =========================
    # Управление доступом
    # =========================
    st.markdown("#### Управление доступом")
    col1, col2 = st.columns(2)
    with col1:
        iam = st.checkbox("IAM (учетные записи)", key="iam")
        iam_v = st.text_input("Производитель IAM", key="iam_v") if iam else ""
        data['Блок 2. IAM'] = iam_v if iam else "Нет"
        
        mfa = st.checkbox("MFA (многофакторная аутентификация)", key="mfa")
        mfa_v = st.text_input("Производитель MFA", key="mfa_v") if mfa else ""
        data['Блок 2. MFA'] = mfa_v if mfa else "Нет"
    with col2:
        pam = st.checkbox("PAM (привилегированный доступ)", key="pam")
        pam_v = st.text_input("Производитель PAM", key="pam_v") if pam else ""
        data['Блок 2. PAM'] = pam_v if pam else "Нет"

    # =========================
    # SOC
    # =========================
    st.markdown("#### Мониторинг и реагирование")
    col1, col2 = st.columns(2)
    with col1:
        siem = st.checkbox("SIEM (мониторинг событий)", key="siem")
        siem_v = st.text_input("Производитель SIEM", key="siem_v") if siem else ""
        data['Блок 2. SIEM'] = siem_v if siem else "Нет"
    with col2:
        soar = st.checkbox("SOAR (автоматизация)", key="soar")
        soar_v = st.text_input("Производитель SOAR", key="soar_v") if soar else ""
        data['Блок 2. SOAR'] = soar_v if soar else "Нет"

    # =========================
    # ДОПОЛНИТЕЛЬНО
    # =========================
    st.markdown("#### Дополнительно")
    col1, col2 = st.columns(2)
    with col1:
        vuln = st.checkbox("Сканер уязвимостей", key="vuln")
        vuln_v = st.text_input("Производитель сканера", key="vuln_v") if vuln else ""
        data['Блок 2. Сканер уязвимостей'] = vuln_v if vuln else "Нет"
    with col2:
        patch = st.checkbox("Patch Management (управление обновлениями)", key="patch")
        patch_v = st.text_input("Производитель Patch Management", key="patch_v") if patch else ""
        data['Блок 2. Patch Management'] = patch_v if patch else "Нет"

    # Валидация
    ib_items = [
        ("EPP", epp, epp_v), ("EDR", edr, edr_v), ("XDR", xdr, xdr_v),
        ("MDR", mdr, mdr_v), ("DLP", dlp, dlp_v), ("Mail Security", mail_sec, mail_v),
        ("CASB", casb, casb_v), ("WAF", waf, waf_v), ("Anti-DDoS", ddos, ddos_v),
        ("IDS/IPS", ids, ids_v), ("NAC", nac, nac_v), ("ZTNA", ztna, ztna_v),
        ("SAST", sast, sast_v), ("DAST", dast, dast_v),
        ("IAM", iam, iam_v), ("MFA", mfa, mfa_v), ("PAM", pam, pam_v),
        ("SIEM", siem, siem_v), ("SOAR", soar, soar_v),
        ("Сканер уязвимостей", vuln, vuln_v),
        ("Patch Management", patch, patch_v),
        ("NAD", nad, nad_v)
    ]
    for name, enabled, vendor in ib_items:
        if enabled and not vendor:
            errors.append(f"Не указан производитель: {name}")

    if errors:
        st.error("Заполните обязательные поля в блоке ИБ:")
        for e in errors:
            st.write(f"- {e}")
        validation_errors.extend(errors)

render_section_feedback(
    "Информационная безопасность",
    get_section_errors(validation_errors, "производитель", "иб"),
    enabled=enable_security
)

# --- БЛОК 3: WEB-РЕСУРСЫ ---
render_section_marker(
    "04 / ЦИФРОВАЯ ПОВЕРХНОСТЬ",
    "Веб-ресурсы",
    "Публичная поверхность, фронтенд-стек и особенности размещения."
)
web_active = st.toggle("Веб-ресурсы", key="web_toggle")
if web_active:
    data['3.1. Хостинг'] = st.selectbox("Хостинг", ["Собственный ЦОД", "Облако KZ", "Облако Global"], key="web_hosting")
    data['3.2. Frontend'] = st.multiselect("Frontend серверы", ["Nginx", "Apache", "IIS", "LiteSpeed", "Cloudflare"], key="web_frontend")
    data['Примечание (Web)'] = st.text_area("Примечания по Web", placeholder="Стек...", key="note_web")

st.divider()

# --- БЛОК 4: РАЗРАБОТКА ---
render_section_marker(
    "05 / РАЗРАБОТКА",
    "Разработка",
    "Команда разработки, языки, CI/CD и дополнительный технологический контекст."
)
dev_count = 0
sel_langs = []
cicd_active = False
dev_active = st.toggle("Разработка", key="dev_toggle")
if dev_active:
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        dev_count = st.number_input("Кол-во разработчиков*", min_value=0, key="dev_cnt_f")
        data['4.1. Разработчики'] = dev_count
        cicd_active = st.checkbox("Используется CI/CD", key="cicd_f")
        data['4.2. CICD'] = "Да" if cicd_active else "Нет"
        if dev_count == 0: validation_errors.append("Укажите количество разработчиков")
    with col_d2:
        lang_list = ["Python", "JavaScript/TypeScript", "Java", "C# / .NET", "PHP", "Go", "C++", "Swift/Kotlin", "Другое"]
        sel_langs = st.multiselect("Языки программирования*", lang_list, key="langs_f")
        if not sel_langs:
            validation_errors.append("Выберите языки разработки")
            data['4.3. Языки разработки'] = "Не указаны"
        elif "Другое" in sel_langs:
            other_l = st.text_input("Укажите другие языки", key="other_langs_f")
            data['4.3. Языки разработки'] = f"{', '.join([l for l in sel_langs if l != 'Другое'])}, {other_l}"
        else:
            data['4.3. Языки разработки'] = ", ".join(sel_langs)
    data['Блок 4. Примечание'] = st.text_area("Примечание к разделу Разработка", placeholder="Стек, фреймворки...", key="note_dev")

render_section_feedback(
    "Разработка",
    get_section_errors(validation_errors, "разработчик", "языки разработки"),
    enabled=dev_active
)
#----Подготовка---
def get_maturity_level(score):
    if score <= 20:
        return "Начальный", "🔴"
    elif score <= 40:
        return "Базовый", "🟠"
    elif score <= 60:
        return "Развивающийся", "🟡"
    elif score <= 80:
        return "Управляемый", "🟢"
    else:
        return "Оптимальный", "🔵"


def calculate_weighted_security_score(enabled, controls):
    if not enabled:
        return 0

    total = 0
    for is_enabled, vendor, weight in controls:
        if is_enabled and str(vendor).strip():
            total += weight

    return min(100, total)


def calculate_it_maturity_score(
    total_arm,
    selected_os_arm,
    sum_os_arm,
    net_active,
    main_speed,
    back_speed,
    selected_routing,
    ap_cnt,
    ngfw_vendor,
    server_active,
    phys_count,
    virt_count,
    selected_virt_sys,
    backup_vendor,
    storage_active,
    st_media_sel,
    cnt_hdd,
    cnt_ssd,
    raid_selected,
    systems_active,
    web_active,
    dev_active,
    dev_count,
    sel_langs,
    cicd_active,
):
    maturity_score = 0

    if total_arm > 0:
        maturity_score += 6
        if selected_os_arm and sum_os_arm == total_arm:
            maturity_score += 5
        if total_arm >= 10:
            maturity_score += 4
        if total_arm >= 50:
            maturity_score += 5
        if total_arm >= 250:
            maturity_score += 5

    if net_active:
        if main_speed > 0:
            maturity_score += 4
        if main_speed >= 100:
            maturity_score += 3
        if back_speed > 0:
            maturity_score += 4
        if selected_routing:
            maturity_score += 4
        if ap_cnt > 0:
            maturity_score += 2
        if str(ngfw_vendor).strip().lower() not in {"", "нет", "no", "none"}:
            maturity_score += 5

    if server_active:
        server_total = phys_count + virt_count
        if phys_count > 0:
            maturity_score += 4
        if virt_count > 0:
            maturity_score += 5
        if server_total >= 5:
            maturity_score += 3
        if selected_virt_sys and virt_count > 0:
            maturity_score += 4
        if str(backup_vendor).strip().lower() not in {"", "нет", "no", "none"}:
            maturity_score += 4

    if storage_active:
        storage_disks = cnt_hdd + cnt_ssd
        if storage_disks > 0:
            maturity_score += 3
        if st_media_sel:
            maturity_score += 3
        if cnt_ssd > 0:
            maturity_score += 2
        if raid_selected:
            maturity_score += 4

    if systems_active:
        maturity_score += 6

    if web_active:
        maturity_score += 4

    if dev_active:
        if dev_count > 0:
            maturity_score += 4
        if dev_count >= 10:
            maturity_score += 3
        if sel_langs:
            maturity_score += 3
        if cicd_active:
            maturity_score += 3

    return min(100, maturity_score)


def build_context(results, client_info):

    users = int(results.get("_user_count", 0))

    industry = str(
        client_info.get(
            "Сфера деятельности",
            ""
        )
    )

    context = {}

    # ==================================
    # COMPANY SIZE
    # ==================================

    context["users"] = users

    context["small_company"] = users < 50

    context["medium_company"] = (
        50 <= users < 250
    )

    context["large_company"] = (
        250 <= users < 1000
    )

    context["enterprise_company"] = (
        users >= 1000
    )

    # ==================================
    # INDUSTRY
    # ==================================

    context["industry"] = industry

    context["is_finance"] = (
        "Финтех" in industry
        or "Банк" in industry
    )

    context["is_gov"] = (
        "Госсектор" in industry
    )

    context["is_it_company"] = (
        "IT" in industry
        or "Разработка" in industry
    )

    context["is_retail"] = (
        "Ритейл" in industry
        or "E-commerce" in industry
    )

    context["is_manufacturing"] = (
        "Производство" in industry
    )

    # ==================================
    # Бизнес-системы
    # ==================================

    result_text = str(results)

    context["has_erp"] = (
        "ERP" in result_text
    )

    context["has_crm"] = (
        "CRM" in result_text
    )

    context["has_accounting"] = (
        "Учет" in result_text
        or "Бухгалтерия" in result_text
    )

    context["has_hrm"] = (
        "HRM" in result_text
    )

    context["has_document_flow"] = (
        "СЭД" in result_text
    )

    context["has_mail"] = (
        results.get(
            "1.5.1. Почтовая система",
            "Нет"
        ) != "Нет"
    )

    # ==================================
    # CRITICAL DATA
    # ==================================

    context["has_personal_data"] = any([
        context["has_hrm"],
        context["has_crm"],
        context["has_accounting"]
    ])

    context["has_critical_systems"] = any([
        context["has_erp"],
        context["has_crm"],
        context["has_accounting"]
    ])

    # ==================================
    # Разработка
    # ==================================

    context["has_development"] = (
        "4.1. Разработчики" in results
    )

    dev_count = int(
        results.get(
            "4.1. Разработчики",
            0
        )
    )

    context["developers"] = dev_count

    context["large_dev_team"] = (
        dev_count >= 10
    )

    # ==================================
    # WEB
    # ==================================

    context["has_public_web"] = (
        "3.1. Хостинг" in results
    )

    # ==================================
    # Инфраструктура
    # ==================================

    context["servers"] = int(
        results.get(
            "Серверы (вирт)",
            0
        )
    ) + int(
        results.get(
            "Серверы (физ)",
            0
        )
    )

    context["has_virtualization"] = (
        context["servers"] > 0
    )

    context["has_backup"] = (
        results.get(
            "Резервное копирование",
            "Нет"
        ) != "Нет"
    )

    # ==================================
    # Безопасность
    # ==================================

    context["has_ngfw"] = (
        results.get(
            "NGFW",
            "Нет"
        ) != "Нет"
    )

    context["has_siem"] = (
        results.get(
            "SIEM",
            "Нет"
        ) != "Нет"
    )

    context["has_edr"] = (
        results.get(
            "EDR",
            "Нет"
        ) != "Нет"
    )

    context["has_patch_management"] = (
        results.get(
            "Patch Management",
            "Нет"
        ) != "Нет"
    )

    context["has_mfa"] = (
        results.get(
            "MFA",
            "Нет"
        ) != "Нет"
    )

    return context

def is_enabled(value):

    if value is None:
        return False

    value = str(value).strip().lower()

    if value in ["", "нет", "none", "false", "-", "n/a"]:
        return False

    return True

def calculate_domain_scores(results):

    domains = {
        "Сетевая безопасность": 0,
        "Защита конечных точек": 0,
        "Идентификация и доступ": 0,
        "Мониторинг и SOC": 0,
        "Резервное копирование": 0,
        "Инфраструктура": 0
    }

    # =========================
    # Сетевая безопасность
    # =========================

    if is_enabled(results.get("NGFW")):
        domains["Сетевая безопасность"] += 25

    if is_enabled(results.get("WAF")):
        domains["Сетевая безопасность"] += 15

    if is_enabled(results.get("Anti-DDoS")):
        domains["Сетевая безопасность"] += 15

    if is_enabled(results.get("VPN")):
        domains["Сетевая безопасность"] += 10

    if is_enabled(results.get("NAC")):
        domains["Сетевая безопасность"] += 20

    if is_enabled(results.get("Сегментация сети")):
        domains["Сетевая безопасность"] += 15

    # =========================
    # Защита конечных устройств
    # =========================

    if is_enabled(results.get("Антивирус")):
        domains["Защита конечных точек"] += 20

    if is_enabled(results.get("EDR")):
        domains["Защита конечных точек"] += 40

    if is_enabled(results.get("Patch Management")):
        domains["Защита конечных точек"] += 20

    if is_enabled(results.get("MDM")):
        domains["Защита конечных точек"] += 10

    if is_enabled(results.get("Device Control")):
        domains["Защита конечных точек"] += 10

    # =========================
    # IAM
    # =========================

    if is_enabled(results.get("MFA")):
        domains["Идентификация и доступ"] += 35

    if is_enabled(results.get("IDM")):
        domains["Идентификация и доступ"] += 25

    if is_enabled(results.get("PAM")):
        domains["Идентификация и доступ"] += 25

    if is_enabled(results.get("SSO")):
        domains["Идентификация и доступ"] += 15

    # =========================
    # MONITORING
    # =========================

    if is_enabled(results.get("SIEM")):
        domains["Мониторинг и SOC"] += 40

    if is_enabled(results.get("SOC")):
        domains["Мониторинг и SOC"] += 30

    if is_enabled(results.get("NAD")):
        domains["Мониторинг и SOC"] += 15

    if is_enabled(results.get("Threat Intelligence")):
        domains["Мониторинг и SOC"] += 15

    # =========================
    # BACKUP
    # =========================

    if is_enabled(results.get("Резервное копирование")):
        domains["Резервное копирование"] += 30

    if is_enabled(results.get("Immutable Backup")):
        domains["Резервное копирование"] += 30

    if is_enabled(results.get("DR")):
        domains["Резервное копирование"] += 20

    if is_enabled(results.get("Air-Gap Backup")):
        domains["Резервное копирование"] += 20

    # =========================
    # INFRA
    # =========================

    if is_enabled(results.get("Виртуализация")):
        domains["Инфраструктура"] += 25

    if is_enabled(results.get("СХД")):
        domains["Инфраструктура"] += 25

    if is_enabled(results.get("Мониторинг")):
        domains["Инфраструктура"] += 20

    if is_enabled(results.get("Резервный канал")):
        domains["Инфраструктура"] += 15

    if is_enabled(results.get("Кластеризация")):
        domains["Инфраструктура"] += 15

    # Ограничение 100%
    for k in domains:
        domains[k] = min(domains[k], 100)

    return domains


def risk_level_label(level):
    labels = {
        "CRITICAL": "Критический",
        "HIGH": "Высокий",
        "MEDIUM": "Средний",
        "LOW": "Низкий",
        "КРИТИЧЕСКИЙ": "Критический",
        "ВЫСОКИЙ": "Высокий",
        "СРЕДНИЙ": "Средний",
        "НИЗКИЙ": "Низкий",
    }
    return labels.get(str(level).upper(), str(level))


# --- Отчет ---
def make_expert_excel(c_info, results, final_score):
    from io import BytesIO
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Аудит ИТ и ИБ"
    # =========================
    # CORPORATE STYLES
    # =========================

    dark_blue_fill = PatternFill(start_color="1F3A5F", end_color="1F3A5F", fill_type="solid")
    light_blue_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    gray_fill = PatternFill(start_color="F3F6F9", end_color="F3F6F9", fill_type="solid")
    critical_fill = PatternFill(start_color="FDE9E7", end_color="FDE9E7", fill_type="solid")
    medium_fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")

    white_font = Font(color="FFFFFF", bold=True)
    section_font = Font(size=14, bold=True, color="1F1F1F")
    normal_font = Font(size=11)

    border = Border(
        left=Side(style='thin', color="D9D9D9"),
        right=Side(style='thin', color="D9D9D9"),
        top=Side(style='thin', color="D9D9D9"),
        bottom=Side(style='thin', color="D9D9D9")
    )
    maturity, maturity_icon = get_maturity_level(final_score)
    domain_scores = calculate_domain_scores(results)

    # =========================
    # EXECUTIVE SUMMARY
    # =========================

    ws.merge_cells('A1:D1')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ ПО КИБЕРБЕЗОПАСНОСТИ"
    ws['A1'].font = Font(bold=True, size=20, color="1F1F1F")

    ws.merge_cells('A3:D3')
    ws['A3'] = f"Компания: {c_info.get('Наименование компании', '-')}"
    ws['A3'].font = Font(bold=True, size=12)

    ws.merge_cells('A4:D4')
    ws['A4'] = f"Дата аудита: {datetime.now().strftime('%d.%m.%Y')}"
    
    ws.merge_cells('A5:D5')
    ws['A5'] = f"{maturity_icon} Уровень зрелости: {maturity}"

    ws.merge_cells('A6:D6')
    ws['A6'] = f"Общая оценка защиты: {final_score}%"

    # Executive Summary Block
    ws.merge_cells('A8:D8')
    ws['A8'] = "УПРАВЛЕНЧЕСКОЕ РЕЗЮМЕ"
    ws['A8'].font = Font(bold=True, size=16, color="FFFFFF")
    ws['A8'].fill = dark_blue_fill
    ws['A8'].alignment = Alignment(horizontal='center')

    summary_text = []

    if results.get("NGFW") != "Нет":
        summary_text.append("• Используется Next-Generation Firewall")

    if results.get("MFA") == "Нет":
        summary_text.append("• Отсутствует многофакторная аутентификация (MFA)")

    if results.get("Блок 2. SIEM") == "Нет":
        summary_text.append("• Отсутствует централизованный мониторинг SIEM")

    if results.get("Резервное копирование") == "Нет":
        summary_text.append("• Не обнаружено централизованное резервное копирование")

    if results.get("_user_count", 0) > 100:
        summary_text.append("• Инфраструктура требует формализации процессов ИБ")

    if not summary_text:
        summary_text.append("• Базовые меры информационной безопасности реализованы")

    ws.merge_cells('A9:D15')
    ws['A9'] = "\n".join(summary_text)
    ws['A9'].alignment = Alignment(wrap_text=True, vertical='top')

    for row in range(9, 16):
        for col in range(1, 5):
            ws.cell(row=row, column=col).fill = gray_fill
            ws.cell(row=row, column=col).border = border

    ws['A9'].font = Font(size=12)

    current_row = 17

    # =========================
    # TOP RISKS OVERVIEW
    # =========================
    context = build_context(results, c_info)
    top_risks = generate_rule_based_risks(
        results,
        context
)
    
    ws.merge_cells(f'A{current_row}:D{current_row}')

    risk_header = ws.cell(
        row=current_row,
        column=1,
        value="КЛЮЧЕВЫЕ КИБЕРРИСКИ"
    )

    risk_header.font = white_font
    risk_header.fill = dark_blue_fill
    risk_header.alignment = Alignment(horizontal='center')

    current_row += 1

    headers = ["#", "Риск", "Критичность", "Влияние на бизнес"]

    for col_num, header in enumerate(headers, 1):

        cell = ws.cell(
            row=current_row,
            column=col_num,
            value=header
        )

        cell.font = white_font
        cell.fill = dark_blue_fill
        cell.border = border

    current_row += 1

    for idx, risk in enumerate(top_risks[:5], start=1):

        level = risk.get("level", "MEDIUM")
        impact = risk.get("impact", "-")

        ws.cell(row=current_row, column=1, value=idx).border = border
        ws.cell(row=current_row, column=2, value=risk.get("risk", "-")).border = border
        ws.cell(row=current_row, column=3, value=risk_level_label(level)).border = border
        ws.cell(row=current_row, column=4, value=impact).border = border

        if "CRITICAL" in str(level).upper():
            fill = critical_fill
        else:
            fill = medium_fill

        for c in range(1, 5):
            ws.cell(row=current_row, column=c).fill = fill

        current_row += 1

    current_row += 2

    # =========================
    # Оценка доменов безопасности
    # =========================

    ws.merge_cells(f'A{current_row}:D{current_row}')
    dom_cell = ws.cell(row=current_row, column=1, value="ОЦЕНКА ДОМЕНОВ БЕЗОПАСНОСТИ")
    dom_cell.font = white_font
    dom_cell.fill = dark_blue_fill
    dom_cell.alignment = Alignment(horizontal='center')

    current_row += 1

    headers = ["Домен безопасности", "Оценка", "Статус"]

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=current_row, column=col_num, value=header)
        cell.font = white_font
        cell.fill = dark_blue_fill
        cell.border = border

    current_row += 1

    for domain, score in domain_scores.items():

        if score >= 80:
            status = "🟢 Сильный"
            fill = light_blue_fill

        elif score >= 50:
            status = "🟠 Средний"
            fill = medium_fill

        else:
            status = "🔴 Слабый"
            fill = critical_fill

        ws.cell(row=current_row, column=1, value=domain).border = border
        ws.cell(row=current_row, column=2, value=f"{score}%").border = border
        ws.cell(row=current_row, column=3, value=status).border = border

        for c in range(1, 4):
            ws.cell(row=current_row, column=c).fill = fill

        current_row += 1

    current_row += 2

    # Основная инфо
    data_info = [
        ("Компания", c_info.get('Наименование компании')),
        ("Сфера", c_info.get('Сфера деятельности')),
        ("Уровень зрелости ИТ/ИБ", f"{final_score}%")
    ]
    curr_row = current_row
    for label, val in data_info:
        ws.cell(row=curr_row, column=1, value=label).border = border
        ws.cell(row=curr_row, column=2, value=val).border = border
        curr_row += 1

    curr_row += 2
    ws.merge_cells(f'A{curr_row}:B{curr_row}')
    ws.cell(row=curr_row, column=1, value="ВЫЯВЛЕННЫЕ РИСКИ И РЕКОМЕНДАЦИИ").font = Font(bold=True, size=14)
    curr_row += 1

    
    # AI Анализ

    context = build_context(
        results,
        c_info
    )
    
    rule_risks = generate_rule_based_risks(
        results,
        context
    )
    
    ai_data = ai_generate_risks_and_recs(
        c_info,
        results
    )
    
    if ai_data:
        ai_data.extend(rule_risks)
    else:
        ai_data = rule_risks
    
    if ai_data:
        for item in ai_data:

            # Уровень и Название
            lvl = item.get('level', 'СРЕДНИЙ')

            ws.merge_cells(f'A{curr_row}:B{curr_row}')

            cell = ws.cell(
                row=curr_row,
                column=1,
                value=f"[{risk_level_label(lvl)}] {item.get('risk', 'Риск')}"
            )

            cell.font = Font(bold=True)

            if "КРИТ" in str(lvl).upper() or "CRITICAL" in str(lvl).upper():
                cell.fill = critical_fill
            else:
                cell.fill = medium_fill

            curr_row += 1

            # Описание, Влияние, Рекомендация
            fields = [
                ("Описание", item.get('description', '-')),
                ("Влияние", item.get('impact', '-')),
                ("Рекомендация", item.get('recommendation', '-')),
                (
                    "Регуляторы",
                    ", ".join(item.get('regulators', []))
                    if isinstance(item.get('regulators'), list)
                    else "-"
                ),
                (
                    "Решения",
                    ", ".join(item.get('vendors', []))
                    if isinstance(item.get('vendors'), list)
                    else "-"
                )
            ]

            for f_label, f_val in fields:

                ws.cell(row=curr_row, column=1, value=f_label).font = Font(italic=True)

                ws.cell(
                    row=curr_row,
                    column=2,
                    value=f_val
                ).alignment = Alignment(wrap_text=True)

                ws.cell(row=curr_row, column=1).border = border
                ws.cell(row=curr_row, column=2).border = border

                curr_row += 1

            curr_row += 1

    # Технические данные
    ws.merge_cells(f'A{curr_row}:D{curr_row}')
    
    sec_cell = ws.cell(row=curr_row, column=1, value="TECHNICAL DETAILS")
    sec_cell.font = white_font
    sec_cell.fill = dark_blue_fill
    sec_cell.alignment = Alignment(horizontal='center')
    curr_row += 1
    for k, v in results.items():
        if not str(k).startswith("_"):
            ws.cell(row=curr_row, column=1, value=k).border = border
            ws.cell(row=curr_row, column=2, value=str(v)).border = border
            if curr_row % 2 == 0:
                ws.cell(row=curr_row, column=1).fill = gray_fill
                ws.cell(row=curr_row, column=2).fill = gray_fill
            curr_row += 1

    ws.column_dimensions['A'].width = 38
    ws.column_dimensions['B'].width = 95
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 20
    
    wb.save(output)
    return output.getvalue()


def build_report_results(
    data,
    main_speed,
    back_speed,
    total_arm,
    ap_cnt,
    selected_routing,
    ngfw_vendor,
    phys_count,
    virt_count,
    v_n_b
):
    results = data.copy()
    results.update({
        "Интернет канал (осн)": f"{main_speed} Mbit/s",
        "Резервный канал": f"{back_speed} Mbit/s",
        "_main_speed": main_speed,
        "_back_speed": back_speed,
        "_user_count": total_arm,
        "WiFi Точки": ap_cnt,
        "WiFi Контроллер": data.get('Wi-Fi Контроллер', "Нет"),
        "Маршрутизация": ", ".join(selected_routing) if selected_routing else "Нет",
        "NGFW": ngfw_vendor if ngfw_vendor else "Нет",
        "Серверы (физ)": phys_count,
        "Серверы (вирт)": virt_count,
        "Резервное копирование": v_n_b if v_n_b else "Нет",
    })

    results["MFA"] = results.get("Блок 2. MFA", "Нет")
    results["SIEM"] = results.get("Блок 2. SIEM", "Нет")
    results["WAF"] = results.get("Блок 2. WAF", "Нет")
    results["Anti-DDoS"] = results.get("Блок 2. Anti-DDoS", "Нет")
    results["EDR"] = results.get("Блок 2. EDR", "Нет")
    results["Patch Management"] = results.get("Блок 2. Patch Management", "Нет")
    return results


def _status_label(score_value):
    if score_value >= 80:
        return "Сильный"
    if score_value >= 50:
        return "Средний"
    return "Слабый"


def _quick_win_signals(results):
    signals = []

    if results.get("MFA") == "Нет":
        signals.append(("critical", "MFA", "Внедрение MFA обычно быстрее всего снижает риск компрометации учетных записей."))

    if results.get("SIEM") == "Нет":
        signals.append(("warn", "SIEM / SOC", "Нет централизованного мониторинга: инциденты сложнее обнаруживать и расследовать."))

    if results.get("EDR") == "Нет":
        signals.append(("warn", "Endpoint", "Без EDR/XDR защита рабочих мест остается слабее против сложных атак."))

    if results.get("Резервное копирование") == "Нет":
        signals.append(("critical", "Backup", "Не указан backup-контур: это критично для устойчивости к ransomware и сбоям."))

    if results.get("Patch Management") == "Нет":
        signals.append(("warn", "Patch Management", "Без централизованных обновлений растет риск эксплуатации известных уязвимостей."))

    legacy_arm = results.get("ОС АРМ (Windows XP/Vista/7/8)", 0)
    legacy_srv = results.get("ОС Сервера (Windows Server 2008/2012 R2)", 0)
    if legacy_arm or legacy_srv:
        signals.append(("critical", "Legacy OS", "Обнаружены устаревшие ОС: нужен план миграции или изоляции."))

    return signals[:5]


def section_status(enabled, complete):
    if not enabled:
        return "disabled"
    return "complete" if complete else "missing"


def calculate_audit_readiness(client_info, validation_errors, section_statuses):
    required_fields = [
        client_info.get('Город'),
        client_info.get('Наименование компании'),
        client_info.get('Сфера деятельности'),
        client_info.get('Сайт компании'),
        client_info.get('Email'),
        client_info.get('ФИО контактного лица'),
        client_info.get('Должность'),
        client_info.get('Контактный телефон'),
    ]

    filled_required = sum(1 for value in required_fields if value)
    required_readiness = (filled_required / len(required_fields)) * 40

    active_sections = [
        status for _, status in section_statuses
        if status != "disabled"
    ]
    completed_sections = sum(
        1 for status in active_sections
        if status == "complete"
    )
    section_readiness = (
        (completed_sections / len(active_sections)) * 45
        if active_sections
        else 0
    )
    validation_readiness = 15 if not validation_errors else 0

    return int(round(
        min(100, required_readiness + section_readiness + validation_readiness)
    ))


def render_audit_cockpit(client_info, results, validation_errors, final_score, it_maturity_score, section_statuses):
    readiness = calculate_audit_readiness(
        client_info,
        validation_errors,
        section_statuses
    )
    domain_scores = calculate_domain_scores(results)
    risk_signals = _quick_win_signals(results)
    maturity, maturity_icon = get_maturity_level(final_score)
    it_maturity, it_maturity_icon = get_maturity_level(it_maturity_score)

    st.sidebar.markdown("### Навигатор аудита")
    st.sidebar.progress(readiness, text=f"Готовность анкеты: {readiness}%")
    st.sidebar.metric("Зрелость ИБ", f"{final_score}%")
    st.sidebar.caption(f"{maturity_icon} {maturity}")
    st.sidebar.metric("Зрелость ИТ", f"{it_maturity_score}%")
    st.sidebar.caption(f"{it_maturity_icon} {it_maturity}")
    st.sidebar.metric("Ошибки заполнения", len(validation_errors))

    st.sidebar.markdown("#### Разделы")
    status_styles = {
        "complete": ("green", "заполнен"),
        "missing": ("red", "не заполнен"),
        "disabled": ("gray", "не включен"),
    }

    for step, status in section_statuses:
        dot_class, title = status_styles.get(status, status_styles["disabled"])
        st.sidebar.markdown(
            f'<div class="sidebar-step" title="{html.escape(title)}">'
            f'<span class="sidebar-dot {dot_class}"></span><span>{html.escape(step)}</span>'
            f'</div>',
            unsafe_allow_html=True
        )

    if validation_errors:
        st.sidebar.markdown("#### Требует внимания")
        for err in list(dict.fromkeys(validation_errors))[:6]:
            st.sidebar.warning(err)


def render_analysis_preview(results, final_score):
    domain_scores = calculate_domain_scores(results)
    risk_signals = _quick_win_signals(results)
    maturity, maturity_icon = get_maturity_level(final_score)

    st.markdown("### Сводка аудита")
    metric_cols = st.columns(4)
    metric_values = [
        ("Зрелость ИБ", f"{final_score}%", f"{maturity_icon} {maturity}"),
        ("Доменов", str(len(domain_scores)), "в текущей оценке"),
        ("Рисков", str(len(risk_signals)), "быстрые сигналы"),
        ("Статус", "Предпросмотр", "перед XLSX"),
    ]

    for col, (label, value, hint) in zip(metric_cols, metric_values):
        with col:
            st.markdown(f"""
            <div class="metric-card">
                <div class="label">{html.escape(label)}</div>
                <div class="value">{html.escape(value)}</div>
                <div class="hint">{html.escape(hint)}</div>
            </div>
            """, unsafe_allow_html=True)

    domain_cards = ''.join(
        f'<div class="domain-card">'
        f'<strong>{html.escape(domain)}</strong>'
        f'<div class="domain-score">{score_value}%</div>'
        f'<div class="hint">{html.escape(_status_label(score_value))}</div>'
        f'</div>'
        for domain, score_value in domain_scores.items()
    )
    st.markdown(f'<div class="domain-row">{domain_cards}</div>', unsafe_allow_html=True)

    if risk_signals:
        st.markdown("#### Быстрые улучшения")
        for kind, title, body in risk_signals:
            st.markdown(f"""
            <div class="risk-chip {html.escape(kind)}">
                <strong>{html.escape(title)}</strong><br>
                {html.escape(body)}
            </div>
            """, unsafe_allow_html=True)


def render_analysis_teaser(results, final_score, validation_errors, section_statuses, client_info):
    domain_scores = calculate_domain_scores(results)
    risk_signals = _quick_win_signals(results)
    maturity, maturity_icon = get_maturity_level(final_score)
    readiness = calculate_audit_readiness(
        client_info,
        validation_errors,
        section_statuses
    )

    weak_domains = sorted(
        domain_scores.items(),
        key=lambda item: item[1]
    )[:3]
    top_signals = risk_signals[:2]

    if top_signals:
        signal_html = "".join(
            f'<div class="mini-signal {html.escape(kind)}">'
            f'<strong>{html.escape(title)}</strong><br>'
            f'{html.escape(body)}'
            f'</div>'
            for kind, title, body in top_signals
        )
    else:
        signal_html = (
            '<div class="mini-signal">'
            '<strong>Быстрые улучшения появятся здесь</strong><br>'
            'Заполните ключевые блоки, и система покажет первые практические рекомендации.'
            '</div>'
        )

    domain_html = "".join(
        f'<div class="mini-domain">'
        f'<span>{html.escape(domain)}</span>'
        f'<strong>{score_value}%</strong>'
        f'</div>'
        for domain, score_value in weak_domains
    )

    st.markdown(f"""
    <div class="analysis-teaser">
        <div class="analysis-teaser-head">
            <div>
                <div class="analysis-teaser-title">Предварительная аналитика уже собирается</div>
                <div class="analysis-teaser-copy">
                    По мере заполнения анкеты здесь появляется ранняя оценка защиты,
                    слабые домены и быстрые улучшения. Полную сводку можно раскрыть ниже.
                </div>
            </div>
            <div class="analysis-pill">обновляется автоматически</div>
        </div>
        <div class="teaser-grid">
            <div class="teaser-card">
                <div class="label">Зрелость ИБ</div>
                <div class="value">{final_score}%</div>
                <div class="hint">{html.escape(maturity_icon)} {html.escape(maturity)}</div>
            </div>
            <div class="teaser-card">
                <div class="label">Готовность анкеты</div>
                <div class="value">{readiness}%</div>
                <div class="hint">до формирования XLSX</div>
            </div>
            <div class="teaser-card">
                <div class="label">Быстрые сигналы</div>
                <div class="value">{len(risk_signals)}</div>
                <div class="hint">первые практические подсказки</div>
            </div>
        </div>
        <div class="teaser-columns">
            <div>
                <div class="label">Что система уже видит</div>
                {signal_html}
            </div>
            <div>
                <div class="label">Домены, требующие внимания</div>
                {domain_html}
            </div>
        </div>
        <div class="analysis-teaser-note">
            Раскройте полный блок ниже, чтобы увидеть все домены, статусы и быстрые улучшения.
        </div>
    </div>
    """, unsafe_allow_html=True)

# --- ИНИЦИАЛИЗАЦИЯ И СТЕК СОСТОЯНИЙ (в самом начале финального блока) ---
if "generation_state" not in st.session_state:
    st.session_state.generation_state = "idle"  # Может быть: idle, preparing, heavy_ai, finalized
if "cached_report_bytes" not in st.session_state:
    st.session_state.cached_report_bytes = None

render_generation_guard(
    st.session_state.generation_state in {"preparing", "heavy_ai"}
)

preview_results = build_report_results(
    data,
    main_speed,
    back_speed,
    total_arm,
    ap_cnt,
    selected_routing,
    ngfw_vendor,
    phys_count,
    virt_count,
    v_n_b
)
audit_started = any(str(value).strip() for value in client_info.values()) or total_arm > 0 or any([
    net_active,
    server_active,
    storage_active,
    is_active,
    enable_security,
    dev_active,
])
security_controls = [
    (epp, epp_v, 8),
    (edr, edr_v, 12),
    (xdr, xdr_v, 8),
    (mdr, mdr_v, 8),
    (dlp, dlp_v, 6),
    (mail_sec, mail_v, 5),
    (casb, casb_v, 4),
    (waf, waf_v, 6),
    (ddos, ddos_v, 4),
    (ids, ids_v, 6),
    (nac, nac_v, 5),
    (ztna, ztna_v, 5),
    (sast, sast_v, 5),
    (dast, dast_v, 5),
    (iam, iam_v, 5),
    (mfa, mfa_v, 10),
    (pam, pam_v, 8),
    (siem, siem_v, 10),
    (soar, soar_v, 4),
    (vuln, vuln_v, 5),
    (patch, patch_v, 6),
    (nad, nad_v, 5),
]
preview_score = calculate_weighted_security_score(enable_security, security_controls)

def has_validation_error(*markers):
    return any(
        any(marker in err for marker in markers)
        for err in validation_errors
    )


general_complete = all([
    client_info.get('Город'),
    client_info.get('Наименование компании'),
    client_info.get('Сфера деятельности'),
    client_info.get('Сайт компании'),
    client_info.get('Email'),
    client_info.get('ФИО контактного лица'),
    client_info.get('Должность'),
    client_info.get('Контактный телефон'),
])
endpoint_complete = total_arm > 0 and bool(selected_os_arm) and sum_os_arm == total_arm
network_complete = net_active and bool(selected_routing) and not has_validation_error(
    "маршрутизации",
    "маршрутизаторов",
    "коммутаторов",
    "Core-уровня",
    "уровня распределения",
    "уровня доступа",
    "Wi-Fi",
    "NGFW",
)
server_complete = server_active and (phys_count > 0 or virt_count > 0) and not has_validation_error(
    "серверов",
    "хостов",
    "резервного копирования",
)
storage_complete = storage_active and not has_validation_error("СХД", "RAID")
systems_complete = is_active and not has_validation_error("название и версию", "версию")
security_complete = enable_security and not has_validation_error("производитель")
web_complete = web_active
dev_complete = dev_active and not has_validation_error("разработчиков", "языки разработки")

it_maturity_score = calculate_it_maturity_score(
    total_arm,
    selected_os_arm,
    sum_os_arm,
    net_active,
    main_speed,
    back_speed,
    selected_routing,
    ap_cnt,
    ngfw_vendor,
    server_active,
    phys_count,
    virt_count,
    selected_virt_sys,
    v_n_b,
    storage_active,
    st_media_sel,
    cnt_hdd,
    cnt_ssd,
    raid_selected,
    is_active,
    web_active,
    dev_active,
    dev_count,
    sel_langs,
    cicd_active,
)

section_statuses = [
    ("Компания", section_status(True, general_complete)),
    ("Конечные точки", section_status(True, endpoint_complete)),
    ("Сеть", section_status(net_active, network_complete)),
    ("Серверы", section_status(server_active, server_complete)),
    ("СХД", section_status(storage_active, storage_complete)),
    ("ИС", section_status(is_active, systems_complete)),
    ("ИБ", section_status(enable_security, security_complete)),
    ("Веб", section_status(web_active, web_complete)),
    ("Разработка", section_status(dev_active, dev_complete)),
]

render_audit_cockpit(
    client_info,
    preview_results,
    validation_errors,
    preview_score,
    it_maturity_score,
    section_statuses
)

render_analysis_teaser(
    preview_results,
    preview_score,
    validation_errors,
    section_statuses,
    client_info
)

with st.expander("Открыть полный предварительный анализ", expanded=False):
    render_analysis_preview(preview_results, preview_score)

# --- ФИНАЛ ---
st.divider()
if validation_errors:
    st.error(f"🚨 Формирование отчета недоступно. Ошибок: {len(validation_errors)}")
    for err in set(validation_errors): st.write(f"- {err}")

# КНОПКА ЗАПУСКА ПРОЦЕССА
# Она активна только тогда, когда процесс еще не запущен
st.markdown("""
Нажимая «Сформировать экспертный отчет», вы даете согласие
на обработку персональных данных в соответствии с
<a href="https://drive.google.com/file/d/1ypEIH9_ePGo3elkR2ifLFBulD5CAFOfs/view?usp=sharing" target="_blank">
Политикой конфиденциальности
</a>.
""", unsafe_allow_html=True)
if st.session_state.generation_state == "idle":
    if st.button("Сформировать экспертный отчет", disabled=len(validation_errors) > 0, type="primary"):
        render_generation_guard(True)
        alert_placeholder = st.empty()
        console_placeholder = st.empty()
        progress_bar = st.progress(0)

        alert_placeholder.markdown(
            """
            <div class="analysis-status-panel">
                <div class="analysis-status-title">Формируется экспертный отчет</div>
                Выполняется нормализация данных, расчет зрелости и сборка XLSX. Это может занять до 4 минут.
                <div class="page-lock-note">Не закрывайте и не обновляйте страницу до завершения формирования.</div>
            </div>
            """,
            unsafe_allow_html=True
        )

        steps = [

            "Инициализация ядра аудита...",
            "Проверка обязательных полей...",
            "Нормализация инфраструктурных данных...",
            "Анализ сетевого периметра...",
            "Анализ защищенности конечных точек...",
            "Проверка устойчивости резервного копирования...",
            "Расчет зрелости киберзащиты...",
            "Построение доменов безопасности...",
            "Глубокий анализ рисков...",
            "Формирование управленческого резюме...",
            "Генерация XLSX отчета...",
            "Финализация артефактов..."
        ]

        active_logs = []

        progress = 0

        for step in steps:

            active_logs.append(
                f"[{time.strftime('%H:%M:%S')}] {step}"
            )

            if len(active_logs) > 4:
                active_logs.pop(0)

            console_placeholder.markdown(
                '<div class="analysis-log">' +
                "".join([f'<div>▶ {line}</div>' for line in active_logs]) +
                '</div>',
                unsafe_allow_html=True
            )

            progress += random.randint(5, 9)

            progress_bar.progress(min(progress, 88))

            time.sleep(random.uniform(0.7, 1.4))


        
        # При клике мы просто меняем статус в памяти на "подготовка" и мгновенно перезапускаем страницу
        st.session_state.generation_state = "preparing"
        st.rerun()

# --- СЦЕНАРИЙ 1: ЭКРАН ОЖИДАНИЯ С ФАКТАМИ ИБ (Показывается СРАЗУ же после клика) ---
if st.session_state.generation_state == "preparing":
    
    # 1. Сразу жестко выводим на экран поле логов и факты информационной безопасности
    st.markdown("#### 🛠️ Ход выполнения анализа:")
    
    # Имитируем лог-систему, как вы просили
    st.info("⚙️ `[СИСТЕМА]`: Инициализация аналитического ядра Khalil Consulting v10.5...")
    st.success("⚙️ `[МАТРИЦА]`: Агрегация параметров ИТ-инфраструктуры успешно завершена.")
    
    st.markdown("---")
    st.markdown("#### Полезные факты и рекомендации по ИБ:")
    
    # Выводим на экран массив фактов в красивом поле, который пользователь будет читать все 3 минуты
    st.markdown("""
    <div class="facts-panel">
        <p><strong>Многофакторная аутентификация:</strong> внедрение MFA блокирует до 99.9% автоматизированных атак на корпоративные учетные записи.</p>
        <p><strong>Защита рабочих мест:</strong> обычного антивируса (EPP) в 2026 году уже недостаточно. Решения класса EDR/XDR критически необходимы для выявления скрытых бесфайловых угроз.</p>
        <p><strong>Безопасность архивов:</strong> резервные копии должны быть изолированы от основной сети. Принцип Immutable Backup снижает риск шифрования бэкапов.</p>
        <p><strong>Сетевой периметр:</strong> сетевая сегментация, VLAN и Zero Trust помогают остановить распространение атаки внутри компании.</p>
        <p><strong>Человеческий фактор:</strong> значительная часть успешных кибератак начинается с фишингового письма. Регулярное обучение снижает этот риск.</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Делаем маленькую паузу в 1.5 секунды, чтобы Streamlit успел железно отправить этот интерфейс в браузер клиента
    time.sleep(1.5)
    
    # Меняем статус на "Запуск тяжелого ИИ" и перезапускаем страницу. 
    # Теперь этот красивый экран останется висеть в браузере, пока ИИ думает!
    st.session_state.generation_state = "heavy_ai"
    st.rerun()

# --- СЦЕНАРИЙ 2: ЗАПУСК ТЯЖЕЛОГО ИИ И СБОРКИ EXCEL ---
if st.session_state.generation_state == "heavy_ai":
    
    # Этот текст и анимация будут гореть параллельно с фактами сверху
    with st.spinner("Производится глубокий анализ рисков..."):
        
        # Подготовка данных перед передачей
        results = build_report_results(
            data,
            main_speed,
            back_speed,
            total_arm,
            ap_cnt,
            selected_routing,
            ngfw_vendor,
            phys_count,
            virt_count,
            v_n_b
        )
        f_score = preview_score
        
        # Запуск функции ИИ (Процессор зависает тут, но на экране пользователя уже горит Сценарий 1 с фактами!)
        report_bytes = make_expert_excel(client_info, results, f_score)
        st.session_state.cached_report_bytes = report_bytes

    # Тихо отправляем в ТГ без создания задержек на экране
    if TOKEN and CHAT_ID:
        try:
            import requests

            telegram_text = (
    "🚨 Новый запрос на аудит!\n"
    f"🏢 Компания: {client_info.get('Наименование компании', '-')}\n"
    f"📍 Город: {client_info.get('Город', '-')}\n"
    f"📊 Сфера: {client_info.get('Сфера деятельности', '-')}\n"
    f"🌐 Сайт: {client_info.get('Сайт компании', '-')}\n"
    f"📧 Email: {client_info.get('Email', '-')}\n"
    f"📞 Телефон: {client_info.get('Контактный телефон', '-')}\n"
    f"👤 Контакт: {client_info.get('ФИО контактного лица', '-')}\n"
    f"💼 Должность: {client_info.get('Должность', '-')}\n"
    f"📊 Уровень зрелости: {f_score}%"
)
            requests.post(f"https://api.telegram.org/bot{TOKEN}/sendMessage", data={"chat_id": CHAT_ID, "text": telegram_text}, timeout=3)
            requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", data={"chat_id": CHAT_ID, "caption": f"Отчет: {client_info['Наименование компании']}"}, files={'document': (f"Audit_v10_{client_info['Наименование компании']}.xlsx", report_bytes)}, timeout=6)
        except Exception:
            pass

    # Переключаем статус в финал
    st.session_state.generation_state = "finalized"
    st.rerun()

# --- СЦЕНАРИЙ 3: ВЫВОД ГОТОВОГО РЕЗУЛЬТАТА ---
if st.session_state.generation_state == "finalized":
    
    st.success("🎉 Экспертный отчет успешно сформирован и проверен системой контроля качества Khalil Consulting!")
    
    st.download_button(
        label="Скачать готовый экспертный отчет (XLSX)",
        data=st.session_state.cached_report_bytes,
        file_name=f"Audit_Khalil_{client_info['Наименование компании']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
    
    # Кнопка для сброса состояния, если пользователь захочет перегенерировать отчет
    if st.button("Сформировать новый отчет"):
        st.session_state.generation_state = "idle"
        st.session_state.cached_report_bytes = None
        st.rerun()

st.info("Khalil Audit System v10.5 | by Ivan Rudoy | Алматы 2026")
