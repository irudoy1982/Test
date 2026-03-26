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

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v3.7")

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

    custom_email_mode = st.checkbox("Email отличается от сайта")
    
    if custom_email_mode:
        client_info['Email'] = st.text_input("Email контактного лица:*", placeholder="info@domain.com")
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
    country_codes = [("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"), ("🇦🇪 +971", "+971")]
    selected_code = p_col1.selectbox("Код", country_codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер телефона", placeholder="777 777 77 77", label_visibility="collapsed")
    client_info['Контактный телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

st.divider()

# --- БЛОК 1: ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ ---
st.header("Блок 1: Информационные технологии")

# 1.1 Конечные точки
st.subheader("1.1. Конечные точки (АРМ)")
data['1.1. Всего АРМ'] = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)
selected_os_arm = st.multiselect("Выберите ОС на АРМ:", ["Windows XP/7/8", "Windows 10", "Windows 11", "Linux", "macOS"], key="ms_arm")
if selected_os_arm:
    for os_item in selected_os_arm:
        data[f"ОС АРМ ({os_item})"] = st.number_input(f"Кол-во АРМ на {os_item}:", min_value=0, step=1, key=f"arm_val_{os_item}")

# 1.2 Сетевая инфраструктура
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_tgl"):
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL/VDSL", "Нет"]
    c_n1, c_n2 = st.columns(2)
    with c_n1:
        m_t = st.selectbox("Основной канал:", net_types, key="main_net_type")
        m_s = st.number_input("Скорость осн. (Mbit/s):", min_value=0, step=10, key="main_net_speed")
        data['1.2.1. Основной канал'] = f"{m_t} ({m_s} Mbps)"
    with c_n2:
        b_t = st.selectbox("Резервный канал:", net_types, key="back_net_type")
        b_s = st.number_input("Скорость рез. (Mbit/s):", min_value=0, step=10, key="back_net_speed")
        data['1.2.2. Резервный канал'] = f"{b_t} ({b_s} Mbps)"

    st.write("**Дополнительные каналы:**")
    cad1, cad2, cad3 = st.columns(3)
    add_ch = []
    if cad1.checkbox("ЕШДИ", key="chk_eshdi"): add_ch.append("ЕШДИ")
    if cad2.checkbox("ЕТСГО", key="chk_etsgo"): add_ch.append("ЕТСГО")
    if cad3.checkbox("VPN", key="chk_vpn"): add_ch.append("VPN")
    data['1.2.3. Доп. каналы'] = ", ".join(add_ch) if add_ch else "Нет"

    st.write("**Активное сетевое оборудование (Уровни):**")
    l1, l2, l3 = st.columns(3)
    with l1:
        if st.checkbox("Ядро (Core)", key="chk_core"):
            data['Сеть: Ядро'] = st.text_input("Вендор (Core):", key="v_core")
    with l2:
        if st.checkbox("Распределение", key="chk_dist"):
            data['Сеть: Распределение'] = st.text_input("Вендор (Dist):", key="v_dist")
    with l3:
        if st.checkbox("Доступ", key="chk_acc"):
            data['Сеть: Доступ'] = st.text_input("Вендор (Access):", key="v_acc")

    if st.checkbox("Межсетевой экран (NGFW)", key="chk_ngfw"):
        v_ng = st.text_input("Производитель NGFW:", key="v_ngfw")
        data['1.2.7. NGFW'] = f"Да ({v_ng if v_ng else 'не указан'})"
        score += 20

# 1.3 Серверы и Виртуализация
st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
cs1, cs2 = st.columns(2)
with cs1:
    data['1.3.1. Физ. серверы'] = st.number_input("Физ. серверы (шт):", min_value=0, step=1, key="phys_srv_ni")
with cs2:
    data['1.3.2. Вирт. серверы'] = st.number_input("Вирт. серверы (шт):", min_value=0, step=1, key="virt_srv_ni")

os_srv = st.multiselect("ОС серверов:", ["Windows Server", "Linux", "Unix"], key="ms_srv")
for o in os_srv:
    data[f"ОС Сервер ({o})"] = st.number_input(f"Кол-во {o}:", min_value=0, step=1, key=f"srv_cnt_{o}")

virt_srv = st.multiselect("Системы виртуализации:", ["VMware", "Hyper-V", "Proxmox", "KVM", "Нет"], key="ms_virt")
if "Нет" not in virt_srv:
    for v in virt_srv:
        data[f"Виртуализация ({v})"] = st.number_input(f"Кол-во хостов {v}:", min_value=0, step=1, key=f"h_cnt_{v}")

if st.checkbox("Резервное копирование", key="chk_backup"):
    v_b = st.text_input("Вендор Бэкапа:", key="v_backup")
    data["Резервное копирование"] = f"Да ({v_b if v_b else 'не указан'})"
    score += 20

# 1.4 СХД (ВОССТАНОВЛЕНО ПОЛНОСТЬЮ)
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Есть СХД", key="st_tgl"):
    data['1.4.1. Типы носителей'] = st.multiselect("Типы носителей в СХД:", ["HDD", "SSD", "NVMe"], key="ms_media")
    data['1.4.2. Конфигурация RAID'] = st.multiselect("Используемые RAID:", ["RAID 1", "RAID 5", "RAID 6", "RAID 10"], key="ms_raid")
    data['1.4.3. Вендор СХД'] = st.text_input("Производитель/Модель СХД:", key="v_storage")
else:
    data['1.4. СХД'] = "Нет"

# 1.5 Информационные системы
st.write("---")
st.subheader("1.5. Внутренние Информационные системы")
if st.toggle("Используются ИС", key="is_tgl"):
    cis1, cis2 = st.columns(2)
    with cis1:
        data['1.5.1. Почта'] = st.selectbox("Почта:", ["Exchange", "M365", "Google", "Yandex", "Свой", "Нет"], key="sb_mail")
        if st.checkbox("Мониторинг", key="chk_mon"):
            data['1.5.2. Мониторинг'] = st.text_input("Система мониторинга:", key="v_mon")
    with cis2:
        st.write("**Информационные системы (РК):**")
        is_list = {"1С": "1c", "Битрикс24": "b24", "Documentolog": "doc", "SAP": "sap", "Directum": "dir", "HelpDesk": "hd"}
        for lab, ks in is_list.items():
            if st.checkbox(lab, key=f"c_{ks}"):
                data[f"ИС: {lab}"] = st.text_input(f"Версия {lab}:", key=f"v_{ks}")

st.divider()

# --- ОСТАЛЬНЫЕ БЛОКИ ---
st.header("Блок 2: Информационная Безопасность")
if st.toggle("Инструменты ИБ", key="ib_tgl"):
    ib_tools = {"DLP": 15, "PAM": 10, "SIEM": 20, "WAF": 10, "EDR": 15}
    for l, p in ib_tools.items():
        if st.checkbox(l, key=f"ib_chk_{l}"):
            v = st.text_input(f"Вендор {l}:", key=f"ib_v_{l}")
            data[l] = f"Да ({v if v else 'не указан'})"
            score += p

st.header("Блок 3: Web-ресурсы")
if st.toggle("Есть свои Web-ресурсы", key="web_tgl"):
    data['3.1. Хостинг'] = st.selectbox("Хостинг:", ["Собственный ЦОД", "Облако KZ", "Облако Global"], key="sb_host")
    data['3.2. Frontend'] = st.multiselect("Frontend:", ["Nginx", "Apache", "Cloudflare", "IIS"], key="ms_front")

st.header("Блок 4: Разработка")
if st.toggle("Своя разработка", key="dev_tgl"):
    data['4.1. Разработчики'] = st.number_input("Кол-во разработчиков:", min_value=0, key="ni_dev")
    data['4.2. CI/CD'] = st.checkbox("Используется CI/CD", key="chk_cicd")

# --- EXCEL И ФИНАЛ ---
def make_excel(c_info, results, f_score):
    out = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.merge_cells('A1:D2')
    ws['A1'] = "ОТЧЕТ ПО АУДИТУ Khalil Trade (2026)"
    ws['A1'].font = Font(bold=True, size=14); ws['A1'].alignment = Alignment(horizontal='center')
    r = 4
    for k, v in {**c_info, "Зрелость": f"{f_score}%"}.items():
        ws.cell(row=r, column=1, value=k).font = Font(bold=True)
        ws.cell(row=r, column=2, value=str(v))
        r
