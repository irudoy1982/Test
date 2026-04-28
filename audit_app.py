import streamlit as st
import pandas as pd
import os
import requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

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

st.title("📋 Опросник Технический аудит ИТ и ИБ (2026) v8.4")

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

    if st.checkbox("Wi-Fi", key="wifi_toggle", help="Наличие корпоративной беспроводной сети."):
        w_col1, w_col2, w_col3 = st.columns(3)
        with w_col1:
            if st.checkbox("Контроллер", key="wifi_ctrl", help="Централизованное управление точками доступа (аппаратное или программное)."):
                wc_v = st.text_input("Производитель/модель контроллера", key="wc_vendor")
                data['Wi-Fi Контроллер'] = wc_v
                if not wc_v: validation_errors.append("Укажите модель Wi-Fi контроллера")
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

if st.checkbox("Резервное копирование", key="ib_backup", help="Наличие специализированного ПО для бэкапа (Veeam, Commvault, Veritas и т.д.)."):
    v_n_b = st.text_input("Вендор Резервного копирования", key="vn_backup", help="Укажите название используемого продукта.")
    data["Резервное копирование"] = v_n_b
    if not v_n_b: validation_errors.append("Укажите вендора резервного копирования")
    score += 20

data['1.3. Примечание'] = st.text_area("Примечание к разделу 1.3", placeholder="Специфика серверного парка...", key="note_1_3")

st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")

if st.toggle("Есть собственная СХД", key="storage_toggle"):
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
if st.toggle("ИС организации", key="is_toggle", help="Бизнес-приложения и корпоративные сервисы."):
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
st.subheader("Информационная безопасность")

col1, col2, col3 = st.columns(3)

errors = []

# --- КОЛОНКА 1 (БАЗОВАЯ ЗАЩИТА) ---
with col1:
    st.markdown("**Базовая защита**")

    siem = st.checkbox("SIEM (система мониторинга и корреляции событий ИБ)", key="siem")
    siem_vendor = st.text_input("Производитель SIEM", key="siem_v") if siem else ""

    dlp = st.checkbox("DLP (предотвращение утечек данных)", key="dlp")
    dlp_vendor = st.text_input("Производитель DLP", key="dlp_v") if dlp else ""

    edr = st.checkbox("EDR (защита рабочих станций)", key="edr")
    edr_vendor = st.text_input("Производитель EDR", key="edr_v") if edr else ""

    email_sec = st.checkbox("Защита почты (антифишинг / антиспам)", key="email_sec")
    email_vendor = st.text_input("Производитель защиты почты", key="email_v") if email_sec else ""


# --- КОЛОНКА 2 (СЕТЬ И ДОСТУП) ---
with col2:
    st.markdown("**Сеть и доступ**")

    waf = st.checkbox("WAF (защита веб-приложений)", key="waf")
    waf_vendor = st.text_input("Производитель WAF", key="waf_v") if waf else ""

    ids = st.checkbox("IDS/IPS (обнаружение и предотвращение атак)", key="ids")
    ids_vendor = st.text_input("Производитель IDS/IPS", key="ids_v") if ids else ""

    nac = st.checkbox("NAC (контроль доступа устройств в сеть)", key="nac")
    nac_vendor = st.text_input("Производитель NAC", key="nac_v") if nac else ""

    ztna = st.checkbox("ZTNA (Zero Trust Network Access)", key="ztna")
    ztna_vendor = st.text_input("Производитель ZTNA", key="ztna_v") if ztna else ""


# --- КОЛОНКА 3 (ADVANCED / SOC) ---
with col3:
    st.markdown("**Расширенная защита**")

    xdr = st.checkbox("XDR (расширенное обнаружение угроз)", key="xdr")
    xdr_vendor = st.text_input("Производитель XDR", key="xdr_v") if xdr else ""

    mdr = st.checkbox("MDR (аутсорс мониторинга ИБ)", key="mdr")
    mdr_vendor = st.text_input("Провайдер MDR", key="mdr_v") if mdr else ""

    soar = st.checkbox("SOAR (автоматизация реагирования)", key="soar")
    soar_vendor = st.text_input("Производитель SOAR", key="soar_v") if soar else ""

    ndr = st.checkbox("NDR (анализ сетевого трафика)", key="ndr")
    ndr_vendor = st.text_input("Производитель NDR", key="ndr_v") if ndr else ""


# --- ДОПОЛНИТЕЛЬНО ---
st.markdown("**Дополнительные меры защиты**")

col4, col5 = st.columns(2)

with col4:
    anti_ddos = st.checkbox("Anti-DDoS (защита от DDoS-атак)", key="ddos")
    ddos_vendor = st.text_input("Производитель Anti-DDoS", key="ddos_v") if anti_ddos else ""

    vuln_scan = st.checkbox("Сканер уязвимостей", key="vuln")
    vuln_vendor = st.text_input("Производитель сканера", key="vuln_v") if vuln_scan else ""

with col5:
    dast = st.checkbox("DAST (динамический анализ безопасности приложений)", key="dast")
    dast_vendor = st.text_input("Производитель DAST", key="dast_v") if dast else ""

    sast = st.checkbox("SAST (статический анализ кода)", key="sast")
    sast_vendor = st.text_input("Производитель SAST", key="sast_v") if sast else ""


# --- ВАЛИДАЦИЯ ---
def check_required(name, enabled, vendor):
    if enabled and not vendor:
        errors.append(f"Не указан производитель для: {name}")

check_required("SIEM", siem, siem_vendor)
check_required("DLP", dlp, dlp_vendor)
check_required("EDR", edr, edr_vendor)
check_required("Защита почты", email_sec, email_vendor)
check_required("WAF", waf, waf_vendor)
check_required("IDS/IPS", ids, ids_vendor)
check_required("NAC", nac, nac_vendor)
check_required("ZTNA", ztna, ztna_vendor)
check_required("XDR", xdr, xdr_vendor)
check_required("MDR", mdr, mdr_vendor)
check_required("SOAR", soar, soar_vendor)
check_required("NDR", ndr, ndr_vendor)
check_required("Anti-DDoS", anti_ddos, ddos_vendor)
check_required("Сканер уязвимостей", vuln_scan, vuln_vendor)
check_required("DAST", dast, dast_vendor)
check_required("SAST", sast, sast_vendor)

if errors:
    st.error("Заполните обязательные поля:")
    for e in errors:
        st.write(f"- {e}")

# --- БЛОК 3: WEB-РЕСУРСЫ ---
st.header("Блок 3: Web-ресурсы")
if st.toggle("Web-ресурсы", key="web_toggle"):
    data['3.1. Хостинг'] = st.selectbox("Хостинг", ["Собственный ЦОД", "Облако KZ", "Облако Global"])
    data['3.2. Frontend'] = st.multiselect("Frontend серверы", ["Nginx", "Apache", "IIS", "LiteSpeed", "Cloudflare"])
    data['Примечание (Web)'] = st.text_area("Примечания по Web", placeholder="Стек...", key="note_web")

st.divider()

# --- БЛОК 4: РАЗРАБОТКА ---
st.header("Блок 4: Разработка")
if st.toggle("Разработка", key="dev_toggle"):
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        dev_count = st.number_input("Кол-во разработчиков*", min_value=0, key="dev_cnt_f")
        data['4.1. Разработчики'] = dev_count
        data['4.2. CICD'] = st.checkbox("Используется CI/CD", key="cicd_f")
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


# --- ГЕНЕРАЦИЯ EXCEL (v8.4) ---
def make_expert_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Отчет ИТ и ИБ"

    # --- СТИЛИ ---
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    bold_font = Font(bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

    def write_block(row, text):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        cell = ws.cell(row=row, column=1, value=text)
        cell.fill = header_fill
        cell.font = white_font
        return row + 1

    def write_kv(row, k, v):
        ws.cell(row=row, column=1, value=k).font = bold_font
        ws.cell(row=row, column=2, value=v)
        return row + 1

    row = 1

    # --- 1. РЕЗЮМЕ ---
    row = write_block(row, "РЕЗЮМЕ ПО ИТ-ИНФРАСТРУКТУРЕ")
    if final_score < 20:
        maturity = "Начальный — процессы отсутствуют, управление реактивное"
    elif final_score < 40:
        maturity = "Базовый — отдельные практики внедрены, нет системности"
    elif final_score < 60:
        maturity = "Управляемый — процессы частично формализованы"
    elif final_score < 80:
        maturity = "Определенный — стандартизированная ИТ-среда"
    else:
        maturity = "Оптимизированный — проактивное управление и автоматизация"

    row = write_kv(row, "Компания", c_info.get("Наименование компании"))
    row = write_kv(row, "Отрасль", c_info.get("Сфера деятельности"))
    row = write_kv(row, "Уровень зрелости", f"{final_score}% — {maturity}")

    verdict = "Инфраструктура содержит критические технологические и организационные риски."
    if final_score > 70:
        verdict = "Инфраструктура в целом устойчива, возможна точечная оптимизация."
    elif final_score < 40:
        verdict = "Высокая вероятность инцидентов и простоев."
    row = write_kv(row, "Общий вывод", verdict)
    row += 1

    # --- 2. КЛЮЧЕВЫЕ РИСКИ ---
    row = write_block(row, "КЛЮЧЕВЫЕ РИСКИ")
    risks = []
    if results.get("Резервное копирование") in [None, "", "Нет"]:
        risks.append("Отсутствие резервного копирования — риск полной остановки бизнеса и потери данных.")
    if results.get("MFA (Аутентификация)") == "Нет":
        risks.append("Отсутствие MFA — высокий риск компрометации учетных записей.")
    if results.get("SIEM (Мониторинг)") == "Нет":
        risks.append("Отсутствие мониторинга — атаки остаются незамеченными.")
    if "0 Mbit" in str(results.get("1.2.2. Резервный канал", "")):
        risks.append("Отсутствие резервного канала — единая точка отказа.")
    raid = str(results.get("1.4.6. RAID-группы", ""))
    if "RAID 0" in raid or "JBOD" in raid:
        risks.append("Использование RAID 0/JBOD — высокий риск потери данных при отказе диска.")
    for i, r in enumerate(risks, 1):
        ws.cell(row=row, column=1, value=f"{i}. {r}")
        row += 1
    row += 1

    # --- 3. ИТ-РИСКИ (CTO VIEW) ---
    row = write_block(row, "ТЕХНОЛОГИЧЕСКИЕ РИСКИ (ИТ)")
    it_risks = []
    virt = results.get("1.3.2. Виртуальные серверы", 0)
    if virt > 50 and results.get("Резервное копирование") in [None, "", "Нет"]:
        it_risks.append("Высокая концентрация сервисов на виртуализации без отказоустойчивости.")
    if results.get("1.1. Всего АРМ", 0) > 100 and results.get("MFA (Аутентификация)") == "Нет":
        it_risks.append("Большое количество пользователей без MFA — масштабируемый риск взлома.")
    for r in it_risks:
        ws.cell(row=row, column=1, value=f"- {r}")
        row += 1
    row += 1

    # --- 4. РЕКОМЕНДАЦИИ ---
    row = write_block(row, "РЕКОМЕНДАЦИИ ПО РАЗВИТИЮ")
    recs = [
        "Организовать резервное копирование по стратегии 3-2-1 с изолированной копией (immutable storage).",
        "Внедрить многофакторную аутентификацию для всех пользователей и администраторов.",
        "Реализовать сегментацию сети (разделение пользователей, серверов, DMZ).",
        "Внедрить централизованный мониторинг (SIEM или аналоги) с хранением логов не менее 90 дней.",
        "Обеспечить отказоустойчивость каналов связи через резервного провайдера.",
        "Оптимизировать СХД: использовать RAID 10/6, внедрить snapshot и replication.",
        "Рассмотреть внедрение EDR/XDR для защиты конечных точек.",
        "Внедрить процессы управления уязвимостями (регулярное сканирование и патчинг)."
    ]
    for r in recs:
        ws.cell(row=row, column=1, value=f"- {r}")
        row += 1
    row += 1

    # --- ВЕНДОРСКИЕ РЕКОМЕНДАЦИИ ---
    row = write_block(row, "РЕКОМЕНДУЕМЫЕ РЕШЕНИЯ И ТЕХНОЛОГИИ")
    vendor_recs = []

    # Проверки с учетом ваших ключей в словаре results (data)
    if "нет" in str(results.get("DLP (Утечки)", "")).lower():
        vendor_recs.append("Защита данных (DLP): Рекомендуется внедрение системы предотвращения утечек данных. Рассмотреть решения: Symantec, Forcepoint, Zecurion, ibatyr, Strac.")
    
    if "нет" in str(results.get("Шифрование", "нет")).lower(): # Ключ 'Шифрование' не задан явно в UI, но оставлен для логики
        vendor_recs.append("Шифрование и защита данных: Для защиты конфиденциальной информации рекомендуется внедрение решений Symantec, Thales, Imperva, Гарда.")

    if "нет" in str(results.get("PAM (Привилегии)", "")).lower():
        vendor_recs.append("Контроль привилегированного доступа (PAM): Рекомендуется внедрение решений CyberArk, WALLIX, Axidian, Netwrix, Fudo или JumpServer.")

    if "нет" in str(results.get("Резервное копирование", "")).lower():
        vendor_recs.append("Резервное копирование: Необходимо внедрение системы резервного копирования уровня enterprise. Рекомендуемые решения: Veeam, Commvault, Veritas.")

    if "нет" in str(results.get("1.2.7. NGFW", "")).lower():
        vendor_recs.append("Сетевая безопасность (NGFW): Рекомендуется внедрение межсетевого экрана нового поколения. Рассмотреть решения: Check Point, Palo Alto, Fortinet, Forcepoint, Cisco, Huawei.")

    if "нет" in str(results.get("IDS/IPS (Атаки)", "")).lower():
        vendor_recs.append("Системы обнаружения и предотвращения атак (IDS/IPS): Рекомендуется внедрение решений Trend Micro или Forcepoint.")

    if "нет" in str(results.get("WAF (Веб)", "")).lower():
        vendor_recs.append("Защита веб-приложений (WAF): Рекомендуется внедрение решений Check Point, Radware, A10, Cloudflare для защиты от атак OWASP Top-10.")

    if "нет" in str(results.get("Anti-DDoS", "")).lower():
        vendor_recs.append("Защита от DDoS: Для обеспечения доступности сервисов рекомендуется использование решений F5, Radware, Check Point.")

    # Корреляционная логика
    if results.get("1.3.2. Виртуальные серверы", 0) > 10 and "нет" in str(results.get("Резервное копирование", "")).lower():
        vendor_recs.append("Виртуальная инфраструктура требует специализированной защиты. Рекомендуется использование решений Veeam или аналогов.")

    if not vendor_recs:
        ws.cell(row=row, column=1, value="Существенных пробелов на уровне выбора технологических решений не выявлено.")
        row += 1
    else:
        for rec in vendor_recs:
            ws.cell(row=row, column=1, value=f"- {rec}")
            row += 1
    row += 1

    # --- 5. ДЕТАЛЬНЫЙ АНАЛИЗ ---
    row = write_block(row, "ДЕТАЛЬНЫЙ АНАЛИЗ")
    headers = ["Параметр", "Значение", "Статус", "Риск", "Рекомендация"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=i, value=h)
        cell.fill = header_fill
        cell.font = white_font
    row += 1

    for k, v in results.items():
        status, risk, rec = "Норма", "Низкий", "Поддерживать текущее состояние"
        val = str(v).lower()
        if "нет" in val or v == 0:
            status, risk, rec = "Проблема", "Средний", "Требуется внедрение или настройка"
        if "резервное копирование" in k.lower() and "нет" in val:
            status, risk, rec = "Критично", "Максимальный", "Необходимо внедрить СРК с регулярным тестированием."
        if "raid" in k.lower() and ("raid 0" in val or "jbod" in val):
            status, risk, rec = "Критично", "Высокий", "Рекомендуется переход на RAID 10 или RAID 6."

        ws.cell(row=row, column=1, value=k)
        ws.cell(row=row, column=2, value=str(v))
        ws.cell(row=row, column=3, value=status)
        ws.cell(row=row, column=4, value=risk)
        ws.cell(row=row, column=5, value=rec)
        row += 1

    widths = {'A': 35, 'B': 25, 'C': 15, 'D': 20, 'E': 60}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    wb.save(output)
    return output.getvalue(), datetime.now().strftime("%d.%m.%Y %H:%M")

# --- ФИНАЛ ---
st.divider()
if validation_errors:
    st.error(f"🚨 Формирование отчета недоступно. Ошибок: {len(validation_errors)}")
    for err in validation_errors: st.write(f"- {err}")

if st.button("📊 Сформировать экспертный отчет", disabled=len(validation_errors) > 0):
    with st.spinner("Обработка..."):
        f_score = min(score, 100)
        report_bytes, final_date = make_expert_excel(client_info, data, f_score)
        try:
            caption = f"🚀 Новый аудит: Khalil Trade\n🏢 {client_info['Наименование компании']}\n📊 Зрелость: {f_score}%"
            requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                            data={"chat_id": CHAT_ID, "caption": caption, "parse_mode": "Markdown"}, 
                            files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report_bytes)})
        except: pass
        st.success("Отчет готов!")
        st.download_button("📥 Скачать отчет", report_bytes, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v8.4 | Almaty 2026")
