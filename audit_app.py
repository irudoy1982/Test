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

TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

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
    client_info['Компания'] = st.text_input("Наименование компании:*")
    site_in = st.text_input("Сайт:*", placeholder="example.kz")
    client_info['Сайт'] = site_in

    if st.checkbox("Email отличается от сайта"):
        client_info['Email'] = st.text_input("Email:*", placeholder="info@domain.com")
    else:
        domain = site_in.replace("https://","").replace("http://","").replace("www.","").split('/')[0]
        if domain and "." in domain:
            st.write("Email (логин):*")
            e1, e2 = st.columns([2, 3])
            prefix = e1.text_input("Login", placeholder="info", label_visibility="collapsed")
            e2.markdown(f"**@{domain}**")
            client_info['Email'] = f"{prefix}@{domain}" if prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    st.write("Телефон:*")
    p1, p2 = st.columns([1, 2])
    code = p1.selectbox("Код", [
        ("🇰🇿 +7","+7"), ("🇷🇺 +7","+7"), ("🇺🇿 +998","+998"), 
        ("🇰🇬 +996","+996"), ("🇹🇯 +992","+992"), ("🇹🇲 +993","+993"),
        ("🇦🇿 +994","+994"), ("🇦🇲 +374","+374"), ("🇧🇾 +375","+375"),
        ("🇬🇪 +995","+995"), ("🇹🇷 +90","+90"), ("🇦🇪 +971","+971"),
        ("🇨🇳 +86","+86"), ("🇺🇸 +1","+1"), ("🇬🇧 +44","+44"), ("🇩🇪 +49","+49")
    ], format_func=lambda x: x[0], label_visibility="collapsed")
    num = p2.text_input("Номер", placeholder="701 123 45 67", label_visibility="collapsed")
    client_info['Телефон'] = f"{code[1]} {num}"

st.divider()

# --- БЛОК 1: ИТ ---
st.header("Блок 1: ИТ Инфраструктура")
data['АРМ (всего)'] = st.number_input("Кол-во АРМ (шт):", min_value=0, step=1)
if st.toggle("Своя сеть"):
    data['Канал'] = st.selectbox("Связь:", ["Оптика", "Радио", "Спутник", "4G/5G"])
    data['NGFW'] = st.text_input("Вендор NGFW:")
    if data['NGFW']: score += 20
data['Серверы (физ)'] = st.number_input("Физ. серверы:", min_value=0)
data['Серверы (вирт)'] = st.number_input("Вирт. серверы:", min_value=0)
data['Почта'] = st.selectbox("Почта:", ["Exchange", "M365", "Google", "Yandex", "Свой", "Нет"])

# --- БЛОК 2: ИБ ---
st.header("Блок 2: Безопасность")
if st.toggle("Системы ИБ"):
    ib_map = {"DLP": 15, "PAM": 10, "SIEM": 20, "WAF": 10, "EDR": 15, "Backup": 20}
    for k, v in ib_map.items():
        if st.checkbox(k):
            vnd = st.text_input(f"Вендор {k}:", key=f"v_{k}")
            data[k] = f"Да ({vnd})" if vnd else "Да"
            score += v
        else: data[k] = "Нет"

# --- EXCEL ---
def get_excel(c_info, res, sc):
    out = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit"
    
    # Стили
    f_white = Font(color="FFFFFF", bold=True)
    f_bold = Font(bold=True)
    fill = PatternFill("solid", start_color="1F4E78")
    al = Alignment(horizontal='center', vertical='center')
    bd = Border(left=Side("thin"), right=Side("thin"), top=Side("thin"), bottom=Side("thin"))

    ws.merge_cells('A1:D2')
    ws['A1'] = "ОТЧЕТ ПО ИТ И ИБ 2026"
    ws['A1'].font = Font(bold=True, size=14, color="1F4E78")
    ws['A1'].alignment = al

    curr = 4
    for k, v in c_info.items():
        ws.cell(curr, 1, k).font = f_bold
        ws.cell(curr, 2, str(v))
        curr += 1
    
    ws.cell(curr+1, 1, "ЗРЕЛОСТЬ ИТ:").font = f_bold
    ws.cell(curr+1, 2, f"{sc}%").font = f_bold
    
    curr += 3
    for i, h in enumerate(["Параметр", "Значение", "Статус", "Рекомендация"], 1):
        c = ws.cell(curr, i, h)
        c.font = f_white; c.fill = fill; c.alignment = al

    curr += 1
    for k, v in res.items():
        ws.cell(curr, 1, k).border = bd
        ws.cell(curr, 2, str(v)).border = bd
        ws.cell(curr, 3, "Проверено").border = bd
        ws.cell(curr, 4, "В норме").border = bd
        curr += 1

    for col, w in {'A': 25, 'B': 25, 'C': 15, 'D': 35}.items():
        ws.column_dimensions[col].width = w
    
    wb.save(out)
    return out.getvalue()

# --- ФИНАЛ ---
st.divider()
if st.button("📊 Сформи
