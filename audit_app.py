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

# Данные для Telegram из Secrets
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# Логотип
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

    if st.checkbox("Email отличается от домена сайта"):
        client_info['Email'] = st.text_input("Email:*", placeholder="info@domain.com")
    else:
        domain = site_in.replace("https://","").replace("http://","").replace("www.","").split('/')[0]
        if domain and "." in domain:
            st.write("Email (логин):*")
            e1, e2 = st.columns([2, 3])
            prefix = e1.text_input("Login", placeholder="info", label_visibility="collapsed", key="em_pre")
            e2.markdown(f"**@{domain}**")
            client_info['Email'] = f"{prefix}@{domain}" if prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    st.write("Телефон:*")
    p1, p2 = st.columns([1, 2])
    codes = [
        ("🇰🇿 +7","+7"), ("🇷🇺 +7","+7"), ("🇺🇿 +998","+998"), 
        ("🇰🇬 +996","+996"), ("🇹🇯 +992","+992"), ("🇦🇪 +971","+971"),
        ("🇹🇷 +90","+90"), ("🇦🇿 +994","+994"), ("🇧🇾 +375","+375"),
        ("🇬🇪 +995","+995"), ("🇺🇸 +1","+1"), ("🇬🇧 +44","+44")
    ]
    code = p1.selectbox("Код", codes, format_func=lambda x: x[0], label_visibility="collapsed")
    num = p2.text_input("Номер", placeholder="701 123 45 67", label_visibility="collapsed", key="ph_num")
    client_info['Телефон'] = f"{code[1]} {num}"

st.divider()

# --- БЛОК 1: ИТ ИНФРАСТРУКТУРА ---
st.header("Блок 1: ИТ Инфраструктура")
data['АРМ (всего)'] = st.number_input("Кол-во АРМ (шт):", min_value=0, step=1)

if st.toggle("Своя сетевая инфраструктура"):
    data['Канал'] = st.selectbox("Тип связи:", ["Оптика", "Радио", "Спутник", "4G/5G", "Starlink"])
    data['NGFW'] = st.text_input("Вендор Межсетевого экрана (NGFW):")
    if data['NGFW']: score += 20

data['Серверы (физ)'] = st.number_input("Физические серверы:", min_value=0)
data['Серверы (вирт)'] = st.number_input("Виртуальные серверы:", min_value=0)
data['Почтовая система'] = st.selectbox("Почта:", ["Exchange", "M365", "Google", "Yandex", "Свой", "Нет"])

# --- БЛОК 2: ИБ ---
st.
