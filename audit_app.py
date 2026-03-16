import streamlit as st
import pandas as pd
import gspread
import os
import socket
from io import BytesIO
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# Excel библиотеки
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage

# --- КОНФИГУРАЦИЯ (Замените на ваши данные) ---
SHEET_NAME = "Audit_Results_2026"
FOLDER_ID = "ВАШ_ID_ПАПКИ" # ID из адресной строки Google Drive

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

# --- 2. ДИАГНОСТИКА СЕТИ (В БОКОВОЙ ПАНЕЛИ) ---
def check_network():
    st.sidebar.subheader("Статус системы")
    try:
        host = "oauth2.google.com"
        ip = socket.gethostbyname(host)
        st.sidebar.success(f"🌐 Связь с Google: OK")
        return True
    except Exception:
        st.sidebar.error("⚠️ Ошибка DNS: Google не доступен")
        return False

is_online = check_network()

# --- 3. ЛОГОТИП И ШАПКА ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### Экспертный аудит вашей ИТ-инфраструктуры")
st.divider()

# --- 4. ИНИЦИАЛИЗАЦИЯ ---
data = {}
client_info = {}
score = 0

# --- 5. ИНФОРМАЦИЯ О КЛИЕНТЕ ---
st.header("🏢 Общая информация")
c1, c2 = st.columns(2)
with c1:
    client_info['Наименование компании'] = st.text_input("Наименование компании")
    client_info['Сайт Компании'] = st.text_input("Сайт Компании")
    client_info['Контактное лицо'] = st.text_input("Контактное лицо (ФИО)")
with c2:
    client_info['Должность'] = st.text_input("Должность")
    client_info['Контактный email'] = st.text_input("Контактный email")
    client_info['Контактный телефон'] = st.text_input("Контактный телефон")

st.divider()

# --- 6. ТЕХНИЧЕСКИЙ БЛОК ---
st.header("📋 Технический аудит")
with st.expander("Инфраструктура и Безопасность", expanded=True):
    data['АРМ'] = st.number_input("Количество АРМ:", min_value=0, step=1)
    data['Серверы'] = st.number_input("Количество серверов:", min_value=0, step=1)
    
    ib_tasks = {
        "Резервное копирование": 20,
        "DLP (Защита от утечек)": 15,
        "EDR/Antimalware": 15,
        "NGFW (Межсетевой экран)": 15
    }
    for task, pts in ib_tasks.items():
        if st.checkbox(task):
            v = st.text_input(f"Вендор {task}:", key=f"v_{task}")
            data[task] = f"Да ({v if v else 'не указан'})"
            score += pts
        else:
            data[task] = "Нет"

# --- 7. ГЕНЕРАЦИЯ EXCEL ---
def make_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit 2026"
    
    # Стили
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    ws.merge_cells('A1:D1')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ 2026"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center')

    ws['A3'] = "КОМПАНИЯ:"; ws['B3'] = c_info['Наименование компании']
    ws['A4'] = "ИНДЕКС ЗРЕЛОСТИ:"; ws['B4'] = f"{final_score}/100"
    
    headers = ["Параметр", "Значение", "Статус", "Рекомендация"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=6, column=i, value=h)
        cell.fill = header_fill; cell.font = white_font

    row = 7
    for k, v in results.items():
        ws.cell(row=row, column=
