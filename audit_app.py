import streamlit as st
import pandas as pd
import requests
import os
from io import BytesIO
from datetime import datetime

# Библиотеки для Excel
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- 1. НАСТРОЙКИ И КОНФИГУРАЦИЯ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

# Берем настройки из Secrets
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ЛОГОТИП И ШАПКА ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### Мы поможем Вам стать лучше!")
st.divider()

# --- 3. СБОР ДАННЫХ КЛИЕНТА ---
st.header("🏢 Информация о компании")
c1, c2 = st.columns(2)
with c1:
    company_name = st.text_input("Наименование компании")
    site = st.text_input("Сайт Компании")
    contact_person = st.text_input("Контактное лицо (ФИО)")
with c2:
    position = st.text_input("Должность")
    email = st.text_input("Контактный email")
    phone = st.text_input("Контактный телефон")

client_info = {
    "Компания": company_name,
    "Лицо": contact_person,
    "Телефон": phone,
    "Email": email
}

st.divider()

# --- 4. ТЕХНИЧЕСКИЙ АУДИТ ---
st.header("📋 Технический аудит")
data = {}
score = 0

with st.expander("Инфраструктура и Безопасность", expanded=True):
    data['АРМ'] = st.number_input("Всего АРМ (шт):", min_value=0, step=1)
    data['Серверы'] = st.number_input("Всего серверов:", min_value=0, step=1)
    
    # Список критических систем ИБ
    ib_tasks = {
        "Резервное копирование": 25,
        "DLP (Защита данных)": 15,
        "EDR/Antimalware": 20,
        "NGFW (Межсетевой экран)": 15,
        "PAM (Контроль доступа)": 15,
        "WAF (Защита сайтов)": 10
    }
    
    for task, pts in ib_tasks.items():
        if st.checkbox(task):
            vendor = st.text_input(f"Вендор {task}:", key=f"v_{task}", placeholder="Например: Kaspersky, Veeam...")
            data[task] = f"Да ({vendor if vendor else 'не указан'})"
            score += pts
        else:
            data[task] = "Нет"

# Логика экспертных рекомендаций
def get_recommendation(key, value):
    db = {
        "Нет": "КРИТИЧЕСКИЙ РИСК. Отсутствие системы снижает устойчивость бизнеса.",
        "Резервное копирование": "КРИТИЧНО: Рекомендуется стратегия 3-2-1. Без бэкапов данные под угрозой.",
        "EDR/Antimalware": "Классических антивирусов недостаточно. Необходим EDR для защиты от шифровальщиков."
    }
    if "Нет" in str(value) or value == 0:
        return db.get(key, db["Нет"])
    return "Конфигурация в норме. Рекомендуется регулярный аудит."

# --- 5. ГЕНЕРАЦИЯ EXCEL ---
def make_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Report"
    
    # Стили
    header_fill = PatternFill(start_color="1F4E78", end_color
