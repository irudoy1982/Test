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
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### Мы поможем Вам стать лучше!")
st.divider()

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v3.2")

data = {}
client_info = {}
score = 0

# --- ШАПКА: ИНФОРМАЦИЯ О КЛИЕНТЕ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город:*")
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    
    site_input = st.text_input("Сайт компании:*", key="site_field", placeholder="example.kz")
    client_info['Сайт компании'] = site_input

    custom_email_mode = st.checkbox("Email отличается от сайта (не рекомендуется)")
    
    if custom_email_mode:
        client_info['Email'] = st.text_input("Email контактного лица:*", placeholder="info@other-domain.com")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица (только логин до @):*")
            e_col1, e_col2 = st.columns([2, 3])
            with e_col1:
                email_prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="email_pre")
            with e_col2:
                st.markdown(f"<div style='padding-top: 5px; font-size: 16px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
            client_info['Email'] = f"{email_prefix}@{clean_domain}" if email_prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    
    st.write("Контактный телефон:*")
    p_col1, p_col2 = st.columns([1, 2])
    country_codes = [
        ("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"),
        ("🇰🇬 +996", "+996"), ("🇹🇯 +992", "+992"), ("🇦🇪 +971", "+971")
    ]
    selected_code = p_col1.selectbox("Код", country_codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер телефона", placeholder="777 777 77 77", label_visibility="collapsed")
    client_info['Контактный телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

st.divider()

# --- БЛОК 1: ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ ---
st.header("Блок 1: Информационные технологии")

# 1.1 Конечные точки
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1, key="total_arm_val")
data['1.1. Всего АРМ'] = total_arm
selected_os_arm = st.multiselect("Выберите ОС на АРМ:", ["Windows XP/Vista/7/8", "Windows 10", "Windows 11", "Linux", "macOS", "Другое"], key="ms_arm_list")
if selected_os_arm:
    for os_item in selected_os_arm:
        count_arm = st.number_input(f"Количество АРМ на {os_item}:", min_value=0, step=1, key=f"arm_cnt_{os_item}")
        data[f"ОС АРМ ({os_item})"] = count_arm

# 1.2 Сетевая инфраструктура
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_toggle"):
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL/VDSL", "Нет"]
    
    col_net1, col_net2 = st.columns(2)
    with col_net1:
        st.write("**Основной канал:**")
        main_type = st.selectbox("Тип (основной):", net_types, key="main_net_type")
        main_speed = st.number_input("Скорость основного (Mbit/s):", min_value=0, step=10, key="main_net_speed")
        data['1.2.1. Основной канал'] = f"{main_type} ({main_speed} Mbit/s)"
        
    with col_net2:
        st.write("**Резервный канал:**")
        back_type = st.selectbox("Тип (резервный):", net_types, key="back_net_type")
        back_speed = st.number_input("Скорость резервного (Mbit/s):", min_value=0, step=10, key="back_net_speed")
        data['1.2.2. Резервный канал'] = f"{back_type} ({back_speed} Mbit/s)"

    st.write("**Дополнительные каналы:**")
    col_add1, col_add2, col_add3 = st.columns(3)
    add_channels = []
    if col_add1.checkbox("ЕШДИ", key="chk_eshdi"): add_channels.append("ЕШДИ")
    if col_add2.checkbox("ЕТСГО", key="chk_etsgo"): add_channels.append("ЕТСГО")
    if col_add3.checkbox("VPN", key="chk_vpn"): add_channels.append("VPN")
    data['1.2.3. Доп. каналы'] = ", ".join(add_channels) if add_channels else "Нет"

    st.write("**Активное оборудование и уровни:**")
    l_col1, l_col2, l_col3 = st.columns(3)
    with l_col1:
        if st.checkbox("Ядро (Core)", key="net_core"):
            core_v = st.text_input("Вендор (Core):", key="core_v")
            data['Сеть: Ядро'] = core_v if core_v else "Да"
    with l_col2:
        if st.checkbox("Распределение", key="net_dist"):
            dist_v = st.text_input("Вендор (Dist):", key="dist_v")
            data['Сеть: Распределение'] = dist_v if dist_v else "Да"
    with l_col3:
        if st.checkbox("Доступ", key="net_acc"):
            acc_v = st.text_input("Вендор (Access):", key="acc_v")
            data['Сеть: Доступ'] = acc_v if acc_v else "Да"

    if st.checkbox("Межсетевой экран (NGFW)", key="ngfw_chk"):
        ngfw_vendor = st.text_input("Производитель (NGFW):", key="ngfw_v")
        data['1.2.7. NGFW'] = f"Да ({ngfw_vendor if ngfw_vendor else 'не указан'})"
        score += 20

# 1.3 Серверы и Виртуализация
st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
col_s1, col_s2 = st.columns(2)
with col_s1:
    data['1.3.1. Физические серверы'] = st.number_input("Физические серверы (шт):", min_value=0, step=1, key="phys_srv")
with col_s2:
    data['1.3.2. Виртуальные серверы'] = st.number_input("Виртуальные серверы (шт):", min_value=0, step=1, key="virt_srv")

if st.checkbox("Резервное копирование", key="ib_backup"):
    v_n_b = st.text_input("Вендор бэкапа:", key="vn_backup")
    data["Резервное копирование"] = f"Да ({v_n_b if v_n_b else 'не указан'})"
    score += 20

# 1.4 СХД
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Есть СХД", key="storage_toggle"):
    data['1.4.1. Типы носителей'] = st.multiselect("Носители:", ["HDD", "SSD", "NVMe"], key="st_media")
    data['1.4.2. RAID'] = st.multiselect("RAID:", ["RAID 1", "RAID 5", "RAID 6", "RAID 10"], key="raid_list")

# 1.5 Внутренние Информационные системы (v3.2)
st.write("---")
st.subheader("1.5. Внутренние Информационные системы")
if st.toggle("Используются внутренние ИС", key="is_block_toggle"):
    col_is1, col_is2 = st.columns(2)
    
    with col_is1:
        st.write("**Почтовая система:**")
        mail_sys = st.selectbox("Тип почты:", ["Exchange (On-Prem)", "Microsoft 365", "Google Workspace", "Yandex Cloud", "Собственный", "Нет"], key="mail_sys")
        data['1.5.1. Почтовая система'] = mail_sys
        
        st.write("**Мониторинг:**")
        has_mon = st.checkbox("Есть мониторинг?", key="mon_chk")
        if has_mon:
            mon_type = st.selectbox("Система:", ["Zabbix", "Nagios", "PRTG", "Prometheus", "Другое"], key="mon_sel")
            data['1.5.2. Мониторинг'] = mon_type

    with col_is2:
        st.write("**Прикладные ИС (Казахстан):**")
        is_kz_list = {
            "1С (Бухгалтерия/ERP)": "is_1c",
            "Битрикс24 (CRM)": "is_b24",
            "Documentolog (СЭД)": "is_doc",
            "SAP": "is_sap",
            "Directum": "is_dir",
            "HelpDesk": "is_hd"
        }
        for label, k_suf in is_kz_list.items():
            if st.checkbox(label, key=f"chk_{k_suf}"):
                ver = st.text_input(f"Версия {label}:", key=f"ver_{k_suf}")
                data[f"ИС: {label}"] = ver if ver else "Да"

st.divider()

# --- БЛОК 2: ИБ ---
st.header("Блок 2: Информационная Безопасность")
if st.toggle("Инструменты ИБ", key="ib_toggle"):
    ib_list = {"DLP": 15, "PAM": 10, "SIEM": 20, "WAF": 10, "EDR": 15}
    for label, pts in ib_list.items():
        if st.checkbox(label, key=f"ib_{label}"):
            v_n = st.text_input(f"Вендор {label}:", key=f"vn_{label}")
            data[label] = f"Да ({v_n if v_n else 'не указан'})"
            score += pts

# --- ФИНАЛ ---
def make_expert_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Report"
    
    ws.merge_cells('A1:D2')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ Khalil Trade (2026)"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center')

    curr_row = 4
    for k, v in c_info.items():
        ws.cell(row=curr_row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=curr_row, column=2, value=str(v))
        curr_row += 1
    
    auto_date = datetime.now().strftime("%d.%m.%Y %H:%M")
    ws.cell(row=curr_row, column=1, value="Зрелость ИТ:").font = Font(bold=True)
    ws.cell(row=curr_row, column=2, value=f"{final_score}%")
    
    curr_row += 2
    headers = ["Параметр", "Значение", "Статус", "Рекомендация"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=curr_row, column=i, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")

    curr_row += 1
    for k, v in results.items():
        ws.cell(row=curr_row, column=1, value=k)
        ws.cell(row=curr_row, column=2, value=str(v))
        ws.cell(row=curr_row, column=3, value="Анализ...")
        curr_row += 1

    wb.save(output)
    return output.getvalue(), auto_date

st.divider()
if st.button("📊 Сформировать отчет", key="btn_final"):
    mandatory = [client_info['Город'], client_info['Наименование компании'], client_info['Сайт компании']]
    if not all(mandatory):
        st.error("Заполните обязательные поля!")
    else:
        with st.spinner("Отправка..."):
            f_score = min(score, 100)
            report_bytes, final_date = make_expert_excel(client_info, data, f_score)
            try:
                caption = f"🚀 *Новый заказ Khalil Trade*\n🏢 {client_info['Наименование компании']}\n📊 Зрелость: {f_score}%"
                requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                              data={"chat_id": CHAT_ID, "caption": caption, "parse_mode": "Markdown"}, 
                              files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report_bytes)})
                st.success("Отправлено!")
                st.balloons()
            except: st.error("Ошибка Telegram")
            st.download_button("📥 Скачать файл", report_bytes, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v3.2 | Almaty 2026")
