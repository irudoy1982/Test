import streamlit as st
import pandas as pd
import os
import requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
# Строки ОБЯЗАТЕЛЬНО должны быть в кавычках
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

# --- НАСТРОЙКИ TELEGRAM ---
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ЛОГОТИП И КОНТАКТЫ ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade IT Audit & Consulting")

st.markdown("### Мы поможем Вам стать лучше!")
st.divider()

st.title("📋 Опросник Технический аудит ИТ и ИБ (2026) v6.2")

# --- ИНСТРУКЦИЯ ДЛЯ ПОЛЬЗОВАТЕЛЯ ---
with st.expander("📖 Инструкция по заполнению (нажмите, чтобы развернуть)"):
    st.markdown("""
    ### Руководство по проведению экспресс-аудита
    
    Данный инструмент предназначен для сбора технических данных об ИТ-ландшафте и уровне защищенности вашей организации.

    1.  **Общая информация**: Укажите корректные контактные данные.
    2.  **Заполнение блоков**: Пройдите по разделам (ИТ, ИБ, Web, Разработка). 
    3.  **Логический контроль**: Следите за тем, чтобы сумма устройств совпадала с общим числом.
    4.  **Результат**: Нажмите кнопку «Сформировать экспертный отчет».
    """)

data = {}
client_info = {}
validation_errors = []
score = 0

# --- ШАПКА ИНФОРМАЦИЯ О КЛИЕНТЕ ---
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

# --- БЛОК 1 ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ ---
st.header("Блок 1 Информационные технологии")

st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт)", min_value=0, step=1)
data['1.1. Всего АРМ'] = total_arm
selected_os_arm = st.multiselect("ОС на АРМ", ["Windows XP/Vista/7/8", "Windows 10", "Windows 11", "Linux", "macOS", "Другое"])

sum_os_arm = 0
if selected_os_arm:
    for os_item in selected_os_arm:
        val = st.number_input(f"Кол-во на {os_item}", min_value=0, step=1, key=f"arm_{os_item}")
        data[f"ОС АРМ ({os_item})"] = val
        sum_os_arm += val

if total_arm > 0 and sum_os_arm != total_arm:
    st.warning(f"⚠️ Ошибка в расчетах АРМ: Указано всего {total_arm}, а по ОС набралось {sum_os_arm}.")
    validation_errors.append("Несовпадение количества АРМ и ОС")

# 1.2 Сетевая инфраструктура
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_toggle"):
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL/VDSL", "Нет"]
    col_net1, col_net2 = st.columns(2)
    with col_net1:
        st.write("Основной канал")
        main_type = st.selectbox("Тип (основной)", net_types, key="main_net_type")
        main_speed = st.number_input("Скорость основного (Mbits)", min_value=0, step=10, key="main_net_speed")
        data['1.2.1. Основной канал'] = f"{main_type} ({main_speed} Mbits)"
    with col_net2:
        st.write("Резервный канал")
        back_type = st.selectbox("Тип (резервный)", net_types, key="back_net_type")
        back_speed = st.number_input("Скорость резервного (Mbits)", min_value=0, step=10, key="back_net_speed")
        data['1.2.2. Резервный канал'] = f"{back_type} ({back_speed} Mbits)"

    st.write("Дополнительные каналы")
    col_add1, col_add2, col_add3 = st.columns(3)
    add_channels = []
    if col_add1.checkbox("ЕШДИ", key="chk_eshdi"): add_channels.append("ЕШДИ")
    if col_add2.checkbox("ЕТСГО", key="chk_etsgo"): add_channels.append("ЕТСГО")
    if col_add3.checkbox("VPN", key="chk_vpn"): add_channels.append("VPN")
    data['1.2.3. Доп. каналы'] = ", ".join(add_channels) if add_channels else "Нет"

    if st.checkbox("Межсетевой экран (NGFW)", key="ngfw_chk"):
        ngfw_vendor = st.text_input("Производитель (NGFW)", key="ngfw_v")
        data['1.2.7. NGFW'] = f"Да ({ngfw_vendor if ngfw_vendor else 'не указан'})"
        score += 20

# 1.3 Серверы
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

if virt_count > 0 and sum_os_srv < virt_count:
    st.warning(f"⚠️ Ошибка: Количество ОС ({sum_os_srv}) меньше количества виртуальных серверов ({virt_count}).")
    validation_errors.append("Недостаточное количество ОС для серверов")

if st.checkbox("Резервное копирование", key="ib_backup"):
    v_n_b = st.text_input("Вендор Резервного копирования", key="vn_backup")
    data["Резервное копирование"] = f"Да ({v_n_b if v_n_b else 'не указан'})"
    score += 20

# --- БЛОК 2 ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ ---
st.header("Блок 2 Информационная Безопасность")
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
    score_cell = ws.cell(row=curr_row, column=2, value=f"{final_score}%")
    bg = "92D050" if final_score > 70 else "FFC000" if final_score > 40 else "FF7C80"
    score_cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
    
    curr_row += 3
    headers = ["Параметр", "Значение", "Статус", "Рекомендация"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=curr_row, column=i, value=h)
        cell.fill = header_fill
        cell.font = white_font

    curr_row += 1
    for k, v in results.items():
        ws.cell(row=curr_row, column=1, value=k).border = border
        ws.cell(row=curr_row, column=2, value=str(v)).border = border
        
        status, rec = "В норме", "Поддерживать состояние."
        is_risk = False
        if "Нет" in str(v) or v == 0:
            is_risk = True; status = "РИСК"; rec = "Рассмотреть внедрение."
        
        st_cell = ws.cell(row=curr_row, column=3, value=status)
        if is_risk: st_cell.font = Font(color="FF0000", bold=True)
        ws.cell(row=curr_row, column=4, value=rec).border = border
        curr_row += 1

    wb.save(output)
    return output.getvalue()

# --- ФИНАЛ ---
st.divider()
if st.button("📊 Сформировать экспертный отчет", disabled=len(validation_errors) > 0):
    mandatory = [client_info['Город'], client_info['Наименование компании'], client_info['ФИО контактного лица'], client_info['Email']]
    if not all(mandatory):
        st.error("⚠️ Заполните все обязательные поля!")
    else:
        with st.spinner("Обработка..."):
            f_score = min(score, 100)
            report_bytes = make_expert_excel(client_info, data, f_score)
            st.success("Отчет готов!")
            st.download_button("📥 Скачать отчет", report_bytes, f"Audit_{client_info['Наименование компании']}.xlsx")
