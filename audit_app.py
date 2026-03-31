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

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v6.2")

# --- ИНСТРУКЦИЯ ДЛЯ ПОЛЬЗОВАТЕЛЯ ---
with st.expander("📖 Инструкция по заполнению (нажмите, чтобы развернуть)"):
    st.markdown("""
    ### Руководство по проведению экспресс-аудита
    
    Данный инструмент предназначен для сбора технических данных об ИТ-ландшафте и уровне защищенности вашей организации.

    1.  **Общая информация**: Укажите корректные контактные данные. Поле Email автоматически подстраивается под домен вашего сайта.
    2.  **Заполнение блоков**: Пройдите по разделам (ИТ, ИБ, Web, Разработка). Используйте переключатели (toggles) для активации нужных подразделов.
    3.  **Логический контроль**: Следите за тем, чтобы сумма устройств с разными ОС совпадала с общим числом АРМ.
    4.  **Результат**: После заполнения всех обязательных полей нажмите кнопку «Сформировать экспертный отчет». 
    """)

data = {}
client_info = {}
validation_errors = []
score = 0

# --- ШАПКА: ИНФОРМАЦИЯ О КЛИЕНТЕ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город")
    client_info['Наименование компании'] = st.text_input("Наименование компании")
    
    site_input = st.text_input("Сайт компании", key="site_field", placeholder="example.kz")
    client_info['Сайт компании'] = site_input

    custom_email_mode = st.checkbox("Email отличается от сайта")
    
    if custom_email_mode:
        client_info['Email'] = st.text_input("Email контактного лица")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица (логин)")
            e_col1, e_col2 = st.columns([2, 3])
            with e_col1:
                email_prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="email_pre")
            with e_col2:
                st.markdown(f"<div style='padding-top: 5px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
            client_info['Email'] = f"{email_prefix}@{clean_domain}" if email_prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица")
    client_info['Должность'] = st.text_input("Должность")
    
    st.write("Контактный телефон")
    p_col1, p_col2 = st.columns([1, 2])
    country_codes = [
        ("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"), ("🇰🇬 +996", "+996"),
        ("🇹🇯 +992", "+992"), ("🇦🇪 +971", "+971"), ("🇹🇷 +90", "+90"), ("🇦🇿 +994", "+994")
    ]
    selected_code = p_col1.selectbox("Код", country_codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed")
    client_info['Контактный телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

st.divider()

# --- БЛОК 1: ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ ---
st.header("Блок 1: Информационные технологии")

st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт)", min_value=0, step=1)
data['1.1. Всего АРМ'] = total_arm
selected_os_arm = st.multiselect("ОС на АРМ", ["Windows XP/Vista/7/8", "Windows 10", "Windows 11", "Linux", "macOS", "Другое"])

sum_os_arm = 0
if selected_os_arm:
    for os_item in selected_os_arm:
        val = st.number_input(f"Кол-во на {os_item}", min_value=0, step=1, key=f"farm_{os_item}")
        data[f"ОС АРМ ({os_item})"] = val
        sum_os_arm += val

if total_arm > 0 and sum_os_arm != total_arm:
    st.warning(f"⚠️ Ошибка: Указано всего {total_arm}, а по ОС набралось {sum_os_arm}.")
    validation_errors.append("Несовпадение количества АРМ и ОС")

st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_toggle"):
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL/VDSL", "Нет"]
    col_net1, col_net2 = st.columns(2)
    with col_net1:
        main_type = st.selectbox("Тип (основной)", net_types, key="main_net_type")
        main_speed = st.number_input("Скорость основного (Mbits)", min_value=0, step=10, key="main_net_speed")
        data['1.2.1. Основной канал'] = f"{main_type} ({main_speed} Mbits)"
    with col_net2:
        back_type = st.selectbox("Тип (резервный)", net_types, key="back_net_type")
        back_speed = st.number_input("Скорость резервного (Mbits)", min_value=0, step=10, key="back_net_speed")
        data['1.2.2. Резервный канал'] = f"{back_type} ({back_speed} Mbits)"

    if st.checkbox("Межсетевой экран (NGFW)", key="ngfw_chk"):
        ngfw_vendor = st.text_input("Производитель (NGFW)", key="ngfw_v")
        data['1.2.7. NGFW'] = f"Да ({ngfw_vendor if ngfw_vendor else 'не указан'})"
        score += 20

st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
col_s1, col_s2 = st.columns(2)
with col_s1:
    phys_count = st.number_input("Количество физических серверов", min_value=0, step=1, key="phys_srv")
    data['1.3.1. Физические серверы'] = phys_count
with col_s2:
    virt_count = st.number_input("Количество виртуальных серверов", min_value=0, step=1, key="virt_srv")
    data['1.3.2. Виртуальные серверы'] = virt_count

s_os_list = ["Windows Server 2008/2012 R2", "Windows Server 2016", "Windows Server 2019", "Windows Server 2022", "Linux", "Unix", "Другое"]
selected_os_srv = st.multiselect("Выберите ОС серверов", s_os_list, key="ms_srv_list")
sum_os_srv = 0
if selected_os_srv:
    for os_s in selected_os_srv:
        val_os = st.number_input(f"Кол-во на {os_s}", min_value=0, key=f"fsrv_{os_s}")
        data[f"ОС Сервера ({os_s})"] = val_os
        sum_os_srv += val_os

if st.checkbox("Резервное копирование", key="ib_backup"):
    v_n_b = st.text_input("Вендор Резервного копирования", key="vn_backup")
    data['Резервное копирование'] = f"Да ({v_n_b if v_n_b else 'не указан'})"
    score += 20

st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Есть собственная СХД", key="storage_toggle"):
    st_media_sel = st.multiselect("Типы носителей", ["HDD (NL-SAS / SATA)", "SSD (SATA / SAS)", "NVMe", "SCM"], key="st_media")
    data['1.4.1. Типы носителей'] = ", ".join(st_media_sel)
    data['1.4.6. RAID-группы'] = st.multiselect("Используемые RAID-группы", ["RAID 0", "RAID 1", "RAID 5", "RAID 6", "RAID 10", "RAID 50", "RAID 60", "JBOD"], key="raid_list")

st.write("---")
st.subheader("1.5. Внутренние Информационные системы")
if st.toggle("ИС организации", key="is_toggle"):
    m_sys = st.selectbox("Почта", ["Exchange (On-Prem)", "Lotus", "Microsoft 365", "Google Workspace", "Собственный", "Нет"])
    data['1.5.1. Почтовая система'] = m_sys
    
    is_list = {"1С": "1c", "Битрикс24": "b24", "Documentolog": "doc", "SAP": "sap"}
    for label, ks in is_list.items():
        if st.checkbox(label):
            data[f"ИС {label}"] = st.text_input(f"Версия {label}", key=f"ver_{ks}")

# --- БЛОК 2: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ ---
st.header("Блок 2: Информационная Безопасность")
if st.toggle("Средства защиты", key="ib_toggle"):
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
            if st.checkbox(label, key=f"fib_{label}"):
                v_n = st.text_input(f"Вендор {label}", key=f"fvn_{label}")
                data[label] = f"Да ({v_n if v_n else 'не указан'})"
                score += pts
            else:
                data[label] = "Нет"

# --- БЛОК 3: WEB-РЕСУРСЫ ---
st.header("Блок 3: Web-ресурсы")
if st.toggle("Web-ресурсы", key="web_toggle"):
    data['3.1. Хостинг'] = st.selectbox("Хостинг", ["Собственный ЦОД", "Облако KZ", "Облако Global"])
    data['3.2. Frontend'] = ", ".join(st.multiselect("Frontend серверы", ["Nginx", "Apache", "IIS", "LiteSpeed", "Cloudflare"]))
    data['Примечание (Web)'] = st.text_area("Примечания по Web", placeholder="Стек...")

# --- БЛОК 4: РАЗРАБОТКА ---
st.header("Блок 4: Разработка")
if st.toggle("Разработка", key="dev_toggle"):
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        dev_count = st.number_input("Кол-во разработчиков", min_value=0)
        data['4.1. Разработчики'] = dev_count
        data['4.2. CI/CD'] = st.checkbox("Используется CI/CD")
    with col_d2:
        lang_list = ["Python", "JavaScript/TypeScript", "Java", "C# / .NET", "PHP", "Go", "C++", "Swift/Kotlin", "Другое"]
        sel_langs = st.multiselect("Языки программирования", lang_list)
        data['4.3. Языки разработки'] = ", ".join(sel_langs) if sel_langs else "Не указаны"

# --- ГЕНЕРАЦИЯ EXCEL ---
def make_expert_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Khalil Audit Report"
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    ws.merge_cells('A1:D2')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ ПО ИТ И ИБ (2026)"
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=16, color="1F4E78")

    curr_row = 4
    for k, v in c_info.items():
        ws.cell(row=curr_row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=curr_row, column=2, value=str(v))
        curr_row += 1
    
    ws.cell(row=curr_row, column=1, value="ИНДЕКС ЗРЕЛОСТИ").font = Font(bold=True)
    ws.cell(row=curr_row, column=2, value=f"{final_score}%").fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    
    curr_row += 3
    headers = ["Параметр", "Значение", "Статус", "Рекомендация"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=curr_row, column=i, value=h)
        cell.fill = header_fill; cell.font = white_font

    curr_row += 1
    for k, v in results.items():
        ws.cell(row=curr_row, column=1, value=k).border = border
        ws.cell(row=curr_row, column=2, value=str(v)).border = border
        ws.cell(row=curr_row, column=3, value="Анализ...").border = border
        ws.cell(row=curr_row, column=4, value="Рекомендация...").border = border
        curr_row += 1

    wb.save(output)
    return output.getvalue()

# --- ФИНАЛ ---
st.divider()
if st.button("📊 Сформировать экспертный отчет", disabled=len(validation_errors) > 0):
    mandatory = [client_info.get('Город'), client_info.get('Наименование компании'), client_info.get('Email')]
    if not all(mandatory):
        st.error("⚠️ Заполните все обязательные поля!")
    else:
        f_score = min(score, 100)
        report_bytes = make_expert_excel(client_info, data, f_score)
        
        # Отправка в Telegram
        try:
            caption = f"🚀 Новый аудит: {client_info['Наименование компании']}\n📊 Зрелость: {f_score}%"
            requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                          data={"chat_id": CHAT_ID, "caption": caption}, 
                          files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report_bytes)})
        except: pass

        st.success("Отчет сформирован!")
        st.download_button("📥 Скачать Excel", report_bytes, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v6.2 | Almaty 2026")
