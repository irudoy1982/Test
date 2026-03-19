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

# Настройки Telegram
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
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    site_input = st.text_input("Сайт компании:*", placeholder="example.kz")
    client_info['Сайт компании'] = site_input

    custom_email_mode = st.checkbox("Email отличается от сайта")
    if custom_email_mode:
        client_info['Email'] = st.text_input("Email контактного лица:*", placeholder="info@domain.com")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица (логин):*")
            e_col1, e_col2 = st.columns([1, 2])
            with e_col1:
                email_prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="email_pre")
            with e_col2:
                st.markdown(f"**@{clean_domain}**")
            client_info['Email'] = f"{email_prefix}@{clean_domain}" if email_prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    client_info['Контактный телефон'] = st.text_input("Контактный телефон:*")

st.divider()

# --- БЛОК 1: ИТ ИНФРАСТРУКТУРА ---
st.header("Блок 1: Информационные технологии")
total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)
data['1.1. Всего АРМ'] = total_arm

if st.toggle("Своя сетевая инфраструктура"):
    data['1.2.1. Тип канала'] = st.selectbox("Тип канала:", ["Оптика", "Радио", "Спутник", "4G/5G", "Starlink"])
    data['1.2.2. NGFW'] = st.text_input("Вендор NGFW:")
    if data['1.2.2. NGFW']: score += 10 # Добавили баллы за периметр

# Серверы
col_s1, col_s2 = st.columns(2)
data['1.3.1. Физ. серверы'] = col_s1.number_input("Физические серверы:", min_value=0, step=1)
data['1.3.2. Вирт. серверы'] = col_s2.number_input("Виртуальные серверы:", min_value=0, step=1)

# Почта и мониторинг
col_p1, col_p2 = st.columns(2)
data['1.4. Почта'] = col_p1.selectbox("Почтовая система:", ["Exchange", "M365", "Google", "Yandex", "Свой сервер", "Нет"])
data['1.5. Мониторинг'] = col_p2.selectbox("Система мониторинга:", ["Zabbix", "PRTG", "Prometheus", "Нет"])
if data['1.5. Мониторинг'] != "Нет": score += 5

st.divider()

# --- БЛОК 2: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ ---
st.header("Блок 2: Информационная Безопасность")
ib_list = {
    "Vulnerability Management (VM/VMDR)": 15, # Твоя специализация
    "DLP (Защита данных)": 15,
    "PAM (Управление доступом)": 10,
    "SIEM (Мониторинг событий)": 15,
    "Backup (Резервное копирование)": 20
}

for label, pts in ib_list.items():
    if st.checkbox(label):
        vendor = st.text_input(f"Укажите вендора для {label}:", key=f"v_{label}")
        data[label] = f"Да ({vendor if vendor else 'не указан'})"
        score += pts
    else:
        data[label] = "Нет"

# --- ГЕНЕРАЦИЯ EXCEL ---
def make_expert_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Khalil Audit Report"
    
    # Стили
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Шапка
    ws.merge_cells('A1:C2')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ ПО ИТ И ИБ (2026)"
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=14, color="1F4E78")

    # Информация о клиенте
    curr_row = 4
    for k, v in c_info.items():
        ws.cell(row=curr_row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=curr_row, column=2, value=str(v))
        curr_row += 1
    
    # Индекс зрелости
    curr_row += 1
    ws.cell(row=curr_row, column=1, value="ИНДЕКС ТЕХНИЧЕСКОЙ ЗРЕЛОСТИ:").font = Font(bold=True)
    res_cell = ws.cell(row=curr_row, column=2, value=f"{final_score}%")
    color = "92D050" if final_score > 70 else "FFC000" if final_score > 40 else "FF7C80"
    res_cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    res_cell.font = Font(bold=True)

    # Таблица результатов
    curr_row += 2
    headers = ["Параметр", "Значение", "Рекомендация"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=curr_row, column=i, value=h)
        cell.fill = header_fill; cell.font = white_font

    curr_row += 1
    for k, v in results.items():
        ws.cell(row=curr_row, column=1, value=k).border = border
        ws.cell(row=curr_row, column=2, value=str(v)).border = border
        rec = "Внедрить" if "Нет" in str(v) else "Оптимизировать"
        ws.cell(row=curr_row, column=3, value=rec).border = border
        curr_row += 1

    for col in ['A', 'B', 'C']: ws.column_dimensions[col].width = 30
    wb.save(output)
    return output.getvalue()

# --- ФИНАЛ ---
if st.button("📊 Сформировать экспертный отчет"):
    if not all([client_info['Компания'], client_info['Email'], client_info['Контактный телефон']]):
        st.error("Заполните все поля со звездочкой!")
    else:
        with st.spinner("Отправка данных..."):
            f_score = min(score, 100)
            report = make_expert_excel(client_info, data, f_score)
            
            # Отправка в TG
            caption = f"🚀 *Новый аудит: {client_info['Наименование компании']}*\n📊 Зрелость: {f_score}%"
            files = {'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report)}
            requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                         data={"chat_id": CHAT_ID, "caption": caption, "parse_mode": "Markdown"}, files=files)
            
            st.success("Отчет отправлен!")
            st.download_button("📥 Скачать Excel", report, "Audit_Khalil.xlsx")
