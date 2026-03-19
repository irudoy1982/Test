import streamlit as st
import pandas as pd
import os
import requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
from datetime import datetime

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

# --- НАСТРОЙКИ TELEGRAM ---
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ЛОГОТИП И ЗАГОЛОВОК ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### Мы поможем Вам стать лучше!")
st.divider()

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026)")

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
            st.write("Email контактного лица (логин):*")
            e_col1, e_col2 = st.columns([2, 3])
            with e_col1:
                email_prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="email_pre")
            with e_col2:
                st.markdown(f"<div style='padding-top: 5px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
            client_info['Email'] = f"{email_prefix}@{clean_domain}" if email_prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    
    st.write("Контактный телефон:*")
    p_col1, p_col2 = st.columns([1, 2])
    with p_col1:
        country_choice = st.selectbox(
            "Код страны",
            options=[
                ("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"), 
                ("🇰🇬 +996", "+996"), ("🇦🇪 +971", "+971"), ("🇹🇷 +90", "+90"),
                ("🇦🇿 +994", "+994"), ("🇦🇲 +374", "+374"), ("🇧🇾 +375", "+375"),
                ("🇬🇪 +995", "+995"), ("🇹🇯 +992", "+992"), ("🇹🇲 +993", "+993"),
                ("🇨🇳 +86", "+86"), ("🇺🇸 +1", "+1"), ("🇬🇧 +44", "+44")
            ],
            format_func=lambda x: x[0],
            label_visibility="collapsed"
        )
    with p_col2:
        phone_val = st.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed")
    
    client_info['Контактный телефон'] = f"{country_choice[1]} {phone_val}"

st.divider()

# --- БЛОК 1: ИТ ИНФРАСТРУКТУРА ---
st.header("Блок 1: Информационные технологии")

data['1.1. Всего АРМ'] = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)

st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_toggle"):
    data['1.2.1. Канал'] = st.selectbox("Тип канала:", ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink"])
    data['1.2.2. NGFW'] = st.text_input("Вендор NGFW:")
    if data['1.2.2. NGFW']: score += 20
else:
    data['1.2. Сетевая инфраструктура'] = "Аренда/Нет"

st.write("---")
col_s1, col_s2 = st.columns(2)
with col_s1:
    data['1.3. Физические серверы'] = st.number_input("Физические серверы:", min_value=0, step=1)
with col_s2:
    data['1.3. Виртуальные серверы'] = st.number_input("Виртуальные серверы:", min_value=0, step=1)

data['1.4. Виртуализация'] = st.multiselect("Системы виртуализации:", ["VMware", "Hyper-V", "Proxmox", "KVM", "Нет"])
data['1.5. Почтовая система'] = st.selectbox("Почта:", ["Exchange", "Microsoft 365", "Google Workspace", "Yandex", "Свой сервер", "Нет"])
data['1.6. Мониторинг'] = st.selectbox("Мониторинг:", ["Нет", "Zabbix", "Nagios", "PRTG", "Prometheus"])

st.divider()

# Блок 2: ИБ
st.header("Блок 2: Информационная Безопасность")
if st.toggle("Есть системы ИБ", key="ib_toggle"):
    ib_list = {"DLP": 15, "PAM": 10, "SIEM": 20, "WAF": 10, "EDR": 15, "Backup": 20}
    for label, pts in ib_list.items():
        if st.checkbox(label):
            vendor = st.text_input(f"Вендор {label}:", key=f"v_{label}")
            data[label] = f"Да ({vendor if vendor else 'не указан'})"
            score += pts
        else:
            data[label] = "Нет"

# --- ГЕНЕРАЦИЯ EXCEL ---
def make_expert_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Report"

    blue_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    ws.merge_cells('A1:D2')
    ws['A1'] = "ОТЧЕТ ПО ИТ И ИБ (2026)"
    ws['A1'].alignment = Alignment(horizontal='center
