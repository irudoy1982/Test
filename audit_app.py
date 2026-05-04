import streamlit as st
import pandas as pd
import os
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

    safe_client = {
        k: v for k, v in c_info.items()
        if k not in forbidden
    }

    safe_results = {
        k: v for k, v in results.items()
        if not any(f.lower() in str(k).lower() for f in forbidden)
    }

    return safe_client, safe_results


def ai_generate_risks_and_recs(c_info, results):
    from openai import OpenAI
    import json

    client = OpenAI()

    safe_client, safe_results = sanitize_for_ai(c_info, results)

    prompt = f"""
Ты выступаешь как CISO и CTO.

Проанализируй ИТ и ИБ состояние компании.

Контекст:
{safe_client}

Данные аудита:
{safe_results}

Требования:
- Учитывай взаимосвязи систем
- Не пиши банальные риски
- Учитывай требования Казахстана (Закон о ПДн, ISO 27001)

Верни JSON:
[
  {{
    "level": "КРИТИЧНО/ВЫСОКИЙ/СРЕДНИЙ",
    "risk": "Название",
    "description": "Описание",
    "impact": "Влияние",
    "recommendation": "Что делать",
    "vendors": ["Vendor1", "Vendor2"]
  }}
]
"""

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}]
    )

    text = response.choices[0].message.content

    try:
        return json.loads(text)
    except:
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

st.title("📋 Опросник Технический аудит ИТ и ИБ (2026) v10.1")

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
def build_context(results, client_info):
    context = {}

    context["industry"] = client_info.get("Сфера деятельности", "")
    context["has_dev"] = "Разработчики" in str(results)
    context["has_critical_systems"] = any(x in str(results) for x in ["ERP", "CRM", "Учет"])
    context["has_personal_data"] = context["has_critical_systems"]
    context["infra_size"] = results.get("_user_count", 0)

    context["is_finance"] = "Финтех" in context["industry"]
    context["is_gov"] = "Госсектор" in context["industry"]

    return context


# --- Отчет ---
def make_expert_excel(c_info, results, final_score):
    from io import BytesIO
    from openpyxl import Workbook
    from openpyxl.styles import Font

    output = BytesIO()
    wb = Workbook()
    ws = wb.active

    row = 1

    # --- AI ДАННЫЕ ---
    ai_data = ai_generate_risks_and_recs(c_info, results)

    # --- EXECUTIVE SUMMARY ---
    ws.cell(row=row, column=1, value="EXECUTIVE SUMMARY").font = Font(bold=True)
    row += 1

    ws.cell(row=row, column=1, value=f"Компания: {c_info.get('Наименование компании')}")
    row += 1

    ws.cell(row=row, column=1, value=f"Уровень зрелости: {final_score}%")
    row += 2

    ws.cell(
        row=row,
        column=1,
        value="В ходе анализа выявлены системные недостатки в архитектуре ИТ и ИБ, которые могут привести к компрометации данных и остановке бизнес-процессов."
    )
    row += 2

    # --- AI АНАЛИЗ ---
    ws.cell(row=row, column=1, value="AI АНАЛИЗ (CISO УРОВЕНЬ)").font = Font(bold=True)
    row += 1

    if not ai_data:
        ws.cell(row=row, column=1, value="AI анализ недоступен")
        row += 2
    else:
        for r in ai_data:
            ws.cell(row=row, column=1, value=f"{r.get('level','')} - {r.get('risk','')}")
            row += 1

            ws.cell(row=row, column=1, value=str(r.get("description", ""))[:800])
            row += 1

            ws.cell(row=row, column=1, value=f"Влияние: {r.get('impact','')}")
            row += 1

            ws.cell(row=row, column=1, value=f"Рекомендация: {r.get('recommendation','')}")
            row += 1

            vendors = ", ".join(r.get("vendors", []))
            ws.cell(row=row, column=1, value=f"Вендоры: {vendors}")
            row += 2

    # --- ДЕТАЛЬНЫЙ АНАЛИЗ ---
    ws.cell(row=row, column=1, value="ДЕТАЛЬНЫЙ АНАЛИЗ").font = Font(bold=True)
    row += 1

    for k, v in results.items():
        if str(k).startswith("_"):
            continue

        ws.cell(row=row, column=1, value=k)
        ws.cell(row=row, column=2, value=str(v))
        row += 1

    wb.save(output)

    return output.getvalue()

# --- ФИНАЛ ---
st.divider()
if validation_errors:
    st.error(f"🚨 Формирование отчета недоступно. Ошибок: {len(validation_errors)}")
    for err in set(validation_errors): st.write(f"- {err}")

if st.button("📊 Сформировать экспертный отчет", disabled=len(validation_errors) > 0):
    with st.spinner("Формирование отчета..."):
        
        # 1. Берем копию всех собранных данных из опросника
        results = data.copy()

        # 2. Добавляем расчетные поля для блока "Резюме" и логики рисков
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

        # Специальная проверка для MFA (поиск в данных, если блок ИБ был активен)
        if "Блок 2. MFA" in results:
            results["MFA"] = results["Блок 2. MFA"]
        else:
            results["MFA"] = "Нет"

        # Расчет итогового балла
        f_score = min(score + 10, 100) 
        
        report_bytes = make_expert_excel(client_info, results, f_score)
        
        try:
            caption = f"🚀 Аудит v10: {client_info['Наименование компании']}\n📊 Зрелость: {f_score}%"
            requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                            data={"chat_id": CHAT_ID, "caption": caption}, 
                            files={'document': (f"Audit_v10_{client_info['Наименование компании']}.xlsx", report_bytes)})
        except: pass
        
        st.success("Отчет успешно сформирован!")
        st.download_button("📥 Скачать экспертный отчет (XLSX)", report_bytes, f"Audit_Khalil_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v10.1 | Production Ready | Almaty 2026")
