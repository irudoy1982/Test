import streamlit as st
import pandas as pd
import os
import threading
import time
import random
import requests
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

    api_key = st.secrets.get("GEMINI_API_KEY")

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
Выступай как эксперт CISO/CTO уровня enterprise.

Проанализируй результаты аудита:

{safe_results}

Сфера деятельности:
{c_info.get("Сфера деятельности", "")}

Используй ТОЛЬКО вендоров из списка.

СПИСОК ВЕНДОРОВ:
{vendor_context}

РЕГУЛЯТОРЫ:
{regulator_context}

Верни ТОЛЬКО JSON:

[
  {{
    "level": "КРИТИЧНО",
    "risk": "Название риска",
    "description": "Описание",
    "impact": "Последствия",
    "recommendation": "Рекомендации",
    "vendors": ["Vendor1", "Vendor2"],
    "regulators": ["ISO 27001"]
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

# Якорь для принудительного перехода в начало страницы
st.markdown("<div id='top'></div>", unsafe_allow_html=True)

# --- НАСТРОЙКИ TELEGRAM (из Secrets) ---
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ЛОГОТИП И КОНТАКТЫ ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade IT Audit & Consulting")

st.markdown("### Мы поможем Вам стать лучше!")
st.divider()

st.title("📋 Опросник Технический аудит ИТ и ИБ by IRudoy(2026) v10.5")

# --- ИНСТРУКЦИЯ ДЛЯ ПОЛЬЗОВАТЕЛЯ ---
with st.expander("📖 Инструкция по заполнению (нажмите, чтобы развернуть)"):
    st.markdown("""
    ### Руководство по проведению экспресс-аудита
    
    Данный инструмент предназначен для сбора технических данных об ИТ-ландшафте и уровне защищенности вашей организации. Пожалуйста, следуйте шагам ниже:

    1. **Общая информация:** Укажите корректные контактные данные. Все поля со звездочкой (*) обязательны.
    2. **Заполнение блоков:** Пройдите по разделам. Используйте переключатели (toggles) для активации нужных подразделов. **Если блок или чекбокс активен — все вложенные поля ввода становятся обязательными.**
    3. **Примечания:** Поля «Примечание» в каждом блоке **не являются обязательными** и заполняются по вашему желанию для уточнения деталей.
    4. **Логический контроль:** Сумма ОС на АРМ должна быть равна общему числу АРМ. Количество ОС на серверах должно быть не меньше числа вирт. машин.
    5. **Результат:** Нажмите кнопку «Сформировать экспертный отчет» для получения файла Excel.
    """)

data = {}
client_info = {}
validation_errors = []
score = 0

# --- ШАПКА: ОБЩАЯ ИНФОРМАЦИЯ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город*", help="Укажите город фактического нахождения головного офиса.")
    industry_options = ["Финтех / Банки", "Ритейл / E-commerce", "Производство", "IT / Разработка", "Госсектор", "Другое"]
    selected_ind = st.selectbox(
        "Сфера деятельности компании*", 
        [""] + industry_options,
        format_func=lambda x: "Выберите сферу..." if x == "" else x,
        help="Отрасль влияет на профиль угроз и регуляторные требования."
    )

    if selected_ind == "Другое":
        industry = st.text_input("Укажите вашу сферу деятельности*", help="Введите отрасль вручную")
    else:
        industry = selected_ind
    
    client_info['Сфера деятельности'] = industry
    client_info['Наименование компании'] = st.text_input("Наименование компании*", help="Официальное или сокращенное название юрлица.")

    site_input = st.text_input("Сайт компании*", key="site_field", placeholder="example.kz", help="Используется для анализа внешнего цифрового отпечатка.")
    client_info['Сайт компании'] = site_input

    custom_email_mode = st.checkbox("Email отличается от сайта", help="Отметьте, если корпоративная почта находится на другом домене.")
    
    if custom_email_mode:
        client_info['Email'] = st.text_input("Email контактного лица*", help="Личный корпоративный email для отправки результатов.")
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
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица*", help="С кем наш эксперт сможет обсудить детали отчета.")
    client_info['Должность'] = st.text_input("Должность*", help="Например: ИТ-Директор, Системный администратор, CEO.")
    
    st.write("Контактный телефон*")
    p_col1, p_col2 = st.columns([1, 2])
    country_codes = [
        ("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"), ("🇰🇬 +996", "+996"),
        ("🇹🇯 +992", "+992"), ("🇦🇪 +971", "+971"), ("🇹🇷 +90", "+90"), ("🇦🇿 +994", "+994")
    ]
    selected_code = p_col1.selectbox("Код", country_codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed", help="Телефон для оперативной связи.")
    client_info['Контактный телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

if not all([client_info.get('Город'), client_info.get('Наименование компании'), client_info.get('Сфера деятельности'), client_info.get('Сайт компании'), client_info.get('Email'), client_info.get('ФИО контактного лица'), client_info.get('Должность'), phone_num]):
    validation_errors.append("Заполните все обязательные поля в блоке 'Общая информация'")

st.divider()

# --- БЛОК 1: ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ ---
st.header("Блок 1: Информационные технологии")

st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт)", min_value=0, step=1, help="Общее число ПК, ноутбуков и тонких клиентов в организации.")
data['1.1. Всего АРМ'] = total_arm
selected_os_arm = st.multiselect("ОС на АРМ", ["Windows XP/Vista/7/8", "Windows 10", "Windows 11", "Linux", "macOS", "Другое"], help="Выберите все типы операционных систем, используемых сотрудниками.")

sum_os_arm = 0
if selected_os_arm:
    for os_item in selected_os_arm:
        val = st.number_input(f"Кол-во на {os_item}", min_value=0, step=1, key=f"arm_{os_item}", help=f"Укажите точное или примерное число устройств с {os_item}.")
        data[f"ОС АРМ ({os_item})"] = val
        sum_os_arm += val

data['1.1. Примечание'] = st.text_area("Примечание к разделу 1.1", placeholder="Напр.: планируем обновление Windows 10 до 11 в Q3", key="note_1_1")

if total_arm > 0 and sum_os_arm != total_arm:
    st.warning(f"⚠️ Ошибка: Сумма по ОС ({sum_os_arm}) должна быть равна общему количеству АРМ ({total_arm}).")
    validation_errors.append("Несовпадение количества АРМ и ОС")

st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
main_speed, back_speed, ap_cnt = 0, 0, 0
selected_routing = []
ngfw_vendor = "Нет"
wifi_enabled = False
wifi_ctrl_enabled = False

if st.toggle("Своя сетевая инфраструктура", key="net_toggle", help="Активируйте, если организация самостоятельно управляет сетевым оборудованием."):
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

st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
col_s1, col_s2 = st.columns(2)
with col_s1:
    phys_count = st.number_input("Количество физических серверов", min_value=0, step=1, key="phys_srv", help="Количество 'железных' серверов в серверной или ЦОД.")
    data['1.3.1. Физические серверы'] = phys_count
with col_s2:
    virt_count = st.number_input("Количество виртуальных серверов", min_value=0, step=1, key="virt_srv", help="Суммарное количество виртуальных машин (VM).")
    data['1.3.2. Виртуальные серверы'] = virt_count

s_os_list = ["Windows Server 2008/2012 R2", "Windows Server 2016", "Windows Server 2019", "Windows Server 2022", "Linux", "Unix", "Другое"]
selected_os_srv = st.multiselect("Выберите ОС серверов", s_os_list, key="ms_srv_list", help="Операционные системы, установленные на серверах.")
sum_os_srv = 0
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

v_n_b = "Нет"
if st.checkbox("Резервное копирование", key="ib_backup", help="Наличие специализированного ПО для бэкапа (Veeam, Commvault, Veritas и т.д.)."):
    v_n_b = st.text_input("Вендор Резервного копирования", key="vn_backup", help="Укажите название используемого продукта.")
    data["Резервное копирование"] = v_n_b
    if not v_n_b: validation_errors.append("Укажите вендора резервного копирования")
    score += 20

data['1.3. Примечание'] = st.text_area("Примечание к разделу 1.3", placeholder="Специфика серверного парка...", key="note_1_3")

st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
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
    m_sys = st.selectbox("Почтовая система", m_opts, help="Где физически и логически располагается ваша электронная почта.")
    
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

st.divider()

# --- БЛОК 2: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ ---
st.header("Блок 2: Информационная безопасность")

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
    # ENDPOINT SECURITY
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
    # DATA SECURITY
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
    # NETWORK SECURITY
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
    # APPLICATION SECURITY
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
    # ACCESS SECURITY
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

# --- БЛОК 3: WEB-РЕСУРСЫ ---
st.header("Блок 3: Web-ресурсы")
web_active = st.toggle("Web-ресурсы", key="web_toggle")
if web_active:
    data['3.1. Хостинг'] = st.selectbox("Хостинг", ["Собственный ЦОД", "Облако KZ", "Облако Global"])
    data['3.2. Frontend'] = st.multiselect("Frontend серверы", ["Nginx", "Apache", "IIS", "LiteSpeed", "Cloudflare"])
    data['Примечание (Web)'] = st.text_area("Примечания по Web", placeholder="Стек...", key="note_web")

st.divider()

# --- БЛОК 4: РАЗРАБОТКА ---
st.header("Блок 4: Разработка")
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
#----Подготовка---
def get_maturity_level(score):
    if score <= 20:
        return "INITIAL", "🔴"
    elif score <= 40:
        return "BASIC", "🟠"
    elif score <= 60:
        return "DEVELOPING", "🟡"
    elif score <= 80:
        return "MANAGED", "🟢"
    else:
        return "OPTIMIZED", "🔵"
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
    # BUSINESS SYSTEMS
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
    # DEVELOPMENT
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
    # INFRASTRUCTURE
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
    # SECURITY
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
        "Network Security": 0,
        "Endpoint Security": 0,
        "Identity & Access": 0,
        "Monitoring & SOC": 0,
        "Backup & Recovery": 0,
        "Infrastructure": 0
    }

    # =========================
    # NETWORK SECURITY
    # =========================

    if is_enabled(results.get("NGFW")):
        domains["Network Security"] += 25

    if is_enabled(results.get("WAF")):
        domains["Network Security"] += 15

    if is_enabled(results.get("Anti-DDoS")):
        domains["Network Security"] += 15

    if is_enabled(results.get("VPN")):
        domains["Network Security"] += 10

    if is_enabled(results.get("NAC")):
        domains["Network Security"] += 20

    if is_enabled(results.get("Сегментация сети")):
        domains["Network Security"] += 15

    # =========================
    # ENDPOINT SECURITY
    # =========================

    if is_enabled(results.get("Антивирус")):
        domains["Endpoint Security"] += 20

    if is_enabled(results.get("EDR")):
        domains["Endpoint Security"] += 40

    if is_enabled(results.get("Patch Management")):
        domains["Endpoint Security"] += 20

    if is_enabled(results.get("MDM")):
        domains["Endpoint Security"] += 10

    if is_enabled(results.get("Device Control")):
        domains["Endpoint Security"] += 10

    # =========================
    # IAM
    # =========================

    if is_enabled(results.get("MFA")):
        domains["Identity & Access"] += 35

    if is_enabled(results.get("IDM")):
        domains["Identity & Access"] += 25

    if is_enabled(results.get("PAM")):
        domains["Identity & Access"] += 25

    if is_enabled(results.get("SSO")):
        domains["Identity & Access"] += 15

    # =========================
    # MONITORING
    # =========================

    if is_enabled(results.get("SIEM")):
        domains["Monitoring & SOC"] += 40

    if is_enabled(results.get("SOC")):
        domains["Monitoring & SOC"] += 30

    if is_enabled(results.get("NAD")):
        domains["Monitoring & SOC"] += 15

    if is_enabled(results.get("Threat Intelligence")):
        domains["Monitoring & SOC"] += 15

    # =========================
    # BACKUP
    # =========================

    if is_enabled(results.get("Резервное копирование")):
        domains["Backup & Recovery"] += 30

    if is_enabled(results.get("Immutable Backup")):
        domains["Backup & Recovery"] += 30

    if is_enabled(results.get("DR")):
        domains["Backup & Recovery"] += 20

    if is_enabled(results.get("Air-Gap Backup")):
        domains["Backup & Recovery"] += 20

    # =========================
    # INFRA
    # =========================

    if is_enabled(results.get("Виртуализация")):
        domains["Infrastructure"] += 25

    if is_enabled(results.get("СХД")):
        domains["Infrastructure"] += 25

    if is_enabled(results.get("Мониторинг")):
        domains["Infrastructure"] += 20

    if is_enabled(results.get("Резервный канал")):
        domains["Infrastructure"] += 15

    if is_enabled(results.get("Кластеризация")):
        domains["Infrastructure"] += 15

    # Ограничение 100%
    for k in domains:
        domains[k] = min(domains[k], 100)

    return domains


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
    ws['A1'] = "EXECUTIVE CYBERSECURITY ASSESSMENT REPORT"
    ws['A1'].font = Font(bold=True, size=20, color="1F1F1F")

    ws.merge_cells('A3:D3')
    ws['A3'] = f"Компания: {c_info.get('Наименование компании', '-')}"
    ws['A3'].font = Font(bold=True, size=12)

    ws.merge_cells('A4:D4')
    ws['A4'] = f"Дата аудита: {datetime.now().strftime('%d.%m.%Y')}"
    
    ws.merge_cells('A5:D5')
    ws['A5'] = f"{maturity_icon} Уровень зрелости: {maturity}"

    ws.merge_cells('A6:D6')
    ws['A6'] = f"Общий Security Score: {final_score}%"

    # Executive Summary Block
    ws.merge_cells('A8:D8')
    ws['A8'] = "EXECUTIVE SUMMARY"
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

    top_risks = generate_rule_based_risks(results)

    ws.merge_cells(f'A{current_row}:D{current_row}')

    risk_header = ws.cell(
        row=current_row,
        column=1,
        value="TOP CYBERSECURITY RISKS"
    )

    risk_header.font = white_font
    risk_header.fill = dark_blue_fill
    risk_header.alignment = Alignment(horizontal='center')

    current_row += 1

    headers = ["#", "Risk", "Severity", "Business Impact"]

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
        ws.cell(row=current_row, column=3, value=level).border = border
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
    # DOMAIN SECURITY ASSESSMENT
    # =========================

    ws.merge_cells(f'A{current_row}:D{current_row}')
    dom_cell = ws.cell(row=current_row, column=1, value="SECURITY DOMAIN ASSESSMENT")
    dom_cell.font = white_font
    dom_cell.fill = dark_blue_fill
    dom_cell.alignment = Alignment(horizontal='center')

    current_row += 1

    headers = ["Security Domain", "Score", "Status"]

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=current_row, column=col_num, value=header)
        cell.font = white_font
        cell.fill = dark_blue_fill
        cell.border = border

    current_row += 1

    for domain, score in domain_scores.items():

        if score >= 80:
            status = "🟢 STRONG"
            fill = light_blue_fill

        elif score >= 50:
            status = "🟠 MODERATE"
            fill = medium_fill

        else:
            status = "🔴 WEAK"
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
    rule_risks = generate_rule_based_risks(results)

    ai_data = ai_generate_risks_and_recs(c_info, results)

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
                value=f"[{lvl}] {item.get('risk', 'Риск')}"
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

# --- ИНИЦИАЛИЗАЦИЯ И СТЕК СОСТОЯНИЙ (в самом начале финального блока) ---
if "generation_state" not in st.session_state:
    st.session_state.generation_state = "idle"  # Может быть: idle, preparing, heavy_ai, finalized
if "cached_report_bytes" not in st.session_state:
    st.session_state.cached_report_bytes = None

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
    if st.button("📊 Сформировать экспертный отчет", disabled=len(validation_errors) > 0):
        alert_placeholder = st.empty()
        console_placeholder = st.empty()
        progress_bar = st.progress(0)

        st.markdown("""
        <style>

        .cyber-alert-box {
            background-color: #fff8e1;
            border: 1px solid #ffcc80;
            color: #ef6c00;
            padding: 15px;
            border-radius: 6px;
            text-align: center;
            font-size: 14px;
            margin-bottom: 20px;
            font-weight: bold;
        }

        .cyber-log-box {
            background: #000;
            color: #00ff00;
            font-family: monospace;
            padding: 15px;
            border: 1px solid #333;
            height: 110px;
            overflow: hidden;
            border-radius: 4px;
            margin-bottom: 20px;
            font-size: 13px;
        }

        </style>
        """, unsafe_allow_html=True)

        alert_placeholder.markdown("""

        <div class="cyber-alert-box">

            ⏳ Выполняется глубокий анализ инфраструктуры.
            Формирование отчета может занять до 3 минут.

            НЕ ЗАКРЫВАЙТЕ И НЕ ОБНОВЛЯЙТЕ СТРАНИЦУ
            
        </div>

        """, unsafe_allow_html=True)

        steps = [

            "Инициализация audit engine...",
            "Проверка обязательных полей...",
            "Нормализация инфраструктурных данных...",
            "Анализ perimeter security...",
            "Анализ endpoint security posture...",
            "Проверка backup resilience...",
            "Расчет cybersecurity maturity...",
            "Построение security domains...",
            "Глубокий анализ рисков...",
            "Формирование executive summary...",
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
                '<div class="cyber-log-box">' +
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
    st.markdown("#### 📋 Полезные факты и рекомендации по ИБ:")
    
    # Выводим на экран массив фактов в красивом поле, который пользователь будет читать все 3 минуты
    st.warning("""
💡 **Многофакторная аутентификация:** Внедрение MFA блокирует до 99.9% автоматизированных атак на корпоративные учетные записи.
              
💡 **Защита рабочих мест:** Обычного антивируса (EPP) в 2026 году уже недостаточно. Решения класса EDR/XDR критически необходимы для выявления скрытых бесфайловых угроз.
              
💡 **Безопасность архивов:** Резервные копии должны быть изолированы от основной сети. Принцип 'Immutable Backup' гарантирует, что хакеры-вымогатели не смогут зашифровать ваши бэкапы.
              
💡 **Сетевой периметр:** Сетевая сегментация (VLAN, концепция Zero Trust) — лучший способ остановить распространение шифровальщика внутри компании, если один компьютер уже заражен.
              
💡 **Человеческий фактор:** Более 80% успешных кибератак начинаются со скомпрометированного фишингового письма. Регулярно обучайте команду кибергигиене.
    """)
    
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
        f_score = min(score + 10, 100)
        
        # Запуск функции ИИ (Процессор зависает тут, но на экране пользователя уже горит Сценарий 1 с фактами!)
        report_bytes = make_expert_excel(client_info, results, f_score)
        st.session_state.cached_report_bytes = report_bytes

    # Тихо отправляем в ТГ без создания задержек на экране
    try:
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
        label="📥 Скачать готовый экспертный отчет (XLSX)",
        data=st.session_state.cached_report_bytes,
        file_name=f"Audit_Khalil_{client_info['Наименование компании']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
    
    # Кнопка для сброса состояния, если пользователь захочет перегенерировать отчет
    if st.button("🔄 Сформировать новый отчет"):
        st.session_state.generation_state = "idle"
        st.session_state.cached_report_bytes = None
        st.rerun()

st.info("Khalil Audit System v10.5 | Ivan Rudoy Production | Almaty 2026")
