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

st.title("📋 Опросник Технический аудит ИТ и ИБ (2026) v8.0")

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
    # ОБЯЗАТЕЛЬНАЯ СФЕРА ДЕЯТЕЛЬНОСТИ
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

# Валидация общих полей
if not all([client_info.get('Город'), client_info.get('Наименование компании'), client_info.get('Сфера деятельности'), client_info.get('Сайт компании'), client_info.get('Email'), client_info.get('ФИО контактного лица'), client_info.get('Должность'), phone_num]):
    validation_errors.append("Заполните все обязательные поля в блоке 'Общая информация'")

st.divider()

# --- БЛОК 1: ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ ---
st.header("Блок 1: Информационные технологии")

# 1.1 Конечные точки
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

# 1.2 Сетевая инфраструктура
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

    # Тип маршрутизации
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

# 1.3 Серверы
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

# 1.4 СХД
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
        cnt_hdd = st.number_input(
            "Количество дисков HDD",
            min_value=0,
            step=1,
            key="cnt_hdd"
        )
        data['1.4.2. Кол-во HDD'] = cnt_hdd

    with col_pct2:
        cnt_ssd = st.number_input(
            "Количество дисков SSD",
            min_value=0,
            step=1,
            key="cnt_ssd"
        )
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
        if allflash:
            score += 5  # небольшой бонус к зрелости

    raid_selected = st.multiselect(
        "Используемые RAID-группы",
        ["RAID 0", "RAID 1", "RAID 5", "RAID 6", "RAID 10", "RAID 50", "RAID 60", "JBOD"],
        key="raid_list"
    )

    data['1.4.6. RAID-группы'] = ", ".join(raid_selected) if raid_selected else "Не указано"

    # Риск-логика (CISO уровень)
    if not raid_selected:
        validation_errors.append("Не указаны RAID-группы СХД")

    if "RAID 0" in raid_selected or "JBOD" in raid_selected:
        score -= 10  # штраф за рискованную конфигурацию

    data['1.4. Примечание'] = st.text_area(
        "Примечание к разделу 1.4",
        placeholder="SAN/NAS, replication, snapshot, DR-site, tiering и т.д.",
        key="note_1_4"
    )

# 1.5 Внутренние ИС
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
        if st.checkbox(label, key=f"is_chk_{ks}", help=f"Используется ли в компании система класса {label}?"):
            c_is1, c_is2 = st.columns(2)
            with c_is1:
                name_is = st.text_input(f"Название продукта {label}*", key=f"name_{ks}", help="Например: 1С, SAP, Bitrix24, Jira.")
            with c_is2:
                ver_is = st.text_input(f"Версия {label}*", key=f"ver_{ks}")
            
            data[f"ИС {label}"] = f"{name_is} (v.{ver_is})"
            if not name_is or not ver_is:
                validation_errors.append(f"Укажите название и версию для {label}")
    
    data['1.5. Примечание'] = st.text_area("Примечание к разделу 1.5", placeholder="Дополнительные ИС...", key="note_1_5")

st.divider()

# --- БЛОК 2: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ ---
st.header("Блок 2: Информационная Безопасность")
if st.toggle("Средства защиты", key="ib_toggle", help="Разверните для указания используемых систем кибербезопасности."):
    ib_systems = {
        "EPP (Антивирус)": 10, "DLP (Утечки)": 15, "PAM (Привилегии)": 10,
        "SIEM (Мониторинг)": 20, "VM (Уязвимости)": 10, "EDR/XDR (Точки)": 15,
        "WAF (Веб)": 10, "Sandbox (Песочница)": 5, "IDS/IPS (Атаки)": 5, "IDM/IGA (Доступ)": 5,
        "MFA (Аутентификация)": 15, "Anti-DDoS": 15
    }
    col_ib1, col_ib2 = st.columns(2)
    items = list(ib_systems.items())
    for i, (label, pts) in enumerate(items):
        target_col = col_ib1 if i < 6 else col_ib2
        with target_col:
            if st.checkbox(label, key=f"fib_{label}", help=f"Наличие внедренного решения класса {label}."):
                v_n = st.text_input(f"Вендор {label}*", key=f"fvn_{label}", help="Например: Касперский, InfoWatch, Positive Technologies и т.д.")
                data[label] = f"Да ({v_n})"
                if not v_n: validation_errors.append(f"Укажите вендора для {label}")
                score += pts
            else:
                data[label] = "Нет"
    
    data['Блок 2. Примечание'] = st.text_area("Примечание к разделу ИБ", placeholder="Планы по внедрению ИБ-решений...", key="note_ib")

st.divider()

# --- БЛОК 3: WEB-РЕСУРСЫ ---
st.header("Блок 3: Web-ресурсы")
if st.toggle("Web-ресурсы", key="web_toggle", help="Анализ публичных сервисов компании."):
    data['3.1. Хостинг'] = st.selectbox("Хостинг", ["Собственный ЦОД", "Облако KZ", "Облако Global"], help="Место размещения веб-сайтов.")
    data['3.2. Frontend'] = st.multiselect("Frontend серверы", ["Nginx", "Apache", "IIS", "LiteSpeed", "Cloudflare"], help="Технологии, отдающие контент пользователям.")
    data['Примечание (Web)'] = st.text_area("Примечания по Web", placeholder="Стек...", key="note_web")

st.divider()

# --- БЛОК 4: РАЗРАБОТКА ---
st.header("Блок 4: Разработка")
if st.toggle("Разработка", key="dev_toggle", help="Заполняется, если в компании есть собственный отдел программирования."):
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        dev_count = st.number_input("Кол-во разработчиков*", min_value=0, key="dev_cnt_f", help="Общий штат внутренних программистов.")
        data['4.1. Разработчики'] = dev_count
        data['4.2. CICD'] = st.checkbox("Используется CI/CD", key="cicd_f", help="Автоматизация сборки и доставки кода (Jenkins, GitLab CI и др.).")
        if dev_count == 0: validation_errors.append("Укажите количество разработчиков")
    with col_d2:
        lang_list = ["Python", "JavaScript/TypeScript", "Java", "C# / .NET", "PHP", "Go", "C++", "Swift/Kotlin", "Другое"]
        sel_langs = st.multiselect("Языки программирования*", lang_list, key="langs_f", help="Основной технологический стек разработки.")
        
        if not sel_langs:
            validation_errors.append("Выберите языки разработки")
            data['4.3. Языки разработки'] = "Не указаны"
        elif "Другое" in sel_langs:
            other_l = st.text_input("Укажите другие языки", key="other_langs_f")
            data['4.3. Языки разработки'] = f"{', '.join([l for l in sel_langs if l != 'Другое'])}, {other_l}"
        else:
            data['4.3. Языки разработки'] = ", ".join(sel_langs)
            
    data['Блок 4. Примечание'] = st.text_area("Примечание к разделу Разработка", placeholder="Стек, фреймворки...", key="note_dev")

# --- ГЕНЕРАЦИЯ EXCEL (v7.0b) ---
def make_expert_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "CISO Report"

    # --- СТИЛИ ---
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    bold_font = Font(bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

    def write_block_title(row, text):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        cell = ws.cell(row=row, column=1, value=text)
        cell.fill = header_fill
        cell.font = white_font
        cell.alignment = Alignment(horizontal='center')
        return row + 1

    def write_kv(row, key, value):
        ws.cell(row=row, column=1, value=key).font = bold_font
        ws.cell(row=row, column=2, value=value)
        return row + 1

    row = 1

    # --- 1. EXECUTIVE SUMMARY ---
    row = write_block_title(row, "EXECUTIVE SUMMARY (CISO VIEW)")

    # Модель зрелости
    if final_score < 20:
        maturity = "Initial"
    elif final_score < 40:
        maturity = "Basic"
    elif final_score < 60:
        maturity = "Managed"
    elif final_score < 80:
        maturity = "Defined"
    else:
        maturity = "Optimized"

    industry = c_info.get("Сфера деятельности", "Не указано")
    total_arm = results.get('1.1. Всего АРМ', 0)

    row = write_kv(row, "Компания", c_info.get("Наименование компании"))
    row = write_kv(row, "Отрасль", industry)
    row = write_kv(row, "Уровень зрелости", f"{maturity} ({final_score}%)")

    verdict = "Инфраструктура частично управляемая, но содержит критические риски."
    if final_score > 70:
        verdict = "Инфраструктура зрелая, требуется точечная оптимизация."
    elif final_score < 40:
        verdict = "Низкий уровень зрелости. Высокая вероятность инцидентов."

    row = write_kv(row, "Вердикт CISO", verdict)

    row += 1

    # --- 2. ТОП РИСКИ ---
    row = write_block_title(row, "TOP-5 CRITICAL RISKS")

    risks = []

    if results.get("Резервное копирование") in [None, "", "Нет"]:
        risks.append("Отсутствует резервное копирование (риск полной потери бизнеса)")

    if results.get("MFA") == "Нет":
        risks.append("Отсутствует MFA (компрометация учетных записей)")

    if "Windows XP/Vista/7/8" in str(results):
        risks.append("Используются устаревшие ОС")

    if results.get("SIEM") == "Нет":
        risks.append("Нет мониторинга безопасности (атаки незаметны)")

    if "0 Mbit" in str(results.get("1.2.2. Резервный канал", "")):
        risks.append("Нет резервного интернет-канала (SPOF)")

    for i, r in enumerate(risks[:5], 1):
        ws.cell(row=row, column=1, value=f"{i}. {r}")
        row += 1

    row += 1

    # --- 3. ATTACK SCENARIO ---
    row = write_block_title(row, "LIKELY ATTACK SCENARIO")

    scenario = [
        "1. Подбор пароля (Password Spraying)",
        "2. Отсутствие MFA → доступ получен",
        "3. Нет SIEM → атака не обнаружена",
        "4. Распространение по сети",
        "5. Шифрование или утечка данных"
    ]

    for step in scenario:
        ws.cell(row=row, column=1, value=step)
        row += 1

    row += 1

    # --- 4. ROADMAP ---
    row = write_block_title(row, "SECURITY ROADMAP")

    roadmap = {
        "0-3 месяца": [
            "Внедрить MFA",
            "Настроить резервное копирование (3-2-1)",
            "Обновить устаревшие ОС"
        ],
        "3-6 месяцев": [
            "Внедрить SIEM",
            "Сегментация сети",
            "Внедрить EDR"
        ],
        "6-12 месяцев": [
            "Zero Trust модель",
            "SOC / мониторинг 24/7",
            "Автоматизация реагирования"
        ]
    }

    for phase, actions in roadmap.items():
        ws.cell(row=row, column=1, value=phase).font = bold_font
        row += 1
        for act in actions:
            ws.cell(row=row, column=2, value=f"- {act}")
            row += 1

    row += 1

    # --- 5. ДЕТАЛЬНЫЙ АНАЛИЗ ---
    row = write_block_title(row, "DETAILED TECHNICAL ANALYSIS")

    headers = ["Параметр", "Значение", "Статус", "Риск", "Рекомендация"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=i, value=h)
        cell.fill = header_fill
        cell.font = white_font

    row += 1

    for k, v in results.items():
        status = "OK"
        risk = "Низкий"
        rec = "Контроль в рамках регламента"

        val_str = str(v).lower()

        if "нет" in val_str or v == 0:
            status = "ПРОБЛЕМА"
            risk = "Средний"
            rec = "Рекомендуется внедрение"

        if "резервное копирование" in k.lower() and ("нет" in val_str):
            status = "КРИТИЧНО"
            risk = "Максимальный"
            rec = "Срочно внедрить backup (3-2-1, immutable)"

        if "mfa" in k.lower() and ("нет" in val_str):
            status = "КРИТИЧНО"
            risk = "Максимальный"
            rec = "Включить MFA для всех пользователей"

        ws.cell(row=row, column=1, value=k)
        ws.cell(row=row, column=2, value=str(v))
        ws.cell(row=row, column=3, value=status)
        ws.cell(row=row, column=4, value=risk)
        ws.cell(row=row, column=5, value=rec)

        row += 1

    # --- ШИРИНА ---
    widths = {'A': 35, 'B': 25, 'C': 15, 'D': 25, 'E': 50}
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

st.info("Khalil Audit System v8.0 | Almaty 2026")
