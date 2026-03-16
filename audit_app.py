import streamlit as st
import pandas as pd
import gspread
import os
from io import BytesIO
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# Excel библиотеки
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage

# --- КОНФИГУРАЦИЯ ---
SHEET_NAME = "Audit_Results_2026"
FOLDER_ID = "1aoHJ5UEw47zh83JyUVi_PE7GzUG90lPw?usp=drive_link" # Вставьте ID вашей папки на Google Диске

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

# --- 2. ЛОГОТИП И КОНТАКТЫ ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### Экспертный анализ вашей инфраструктуры")
st.divider()

# --- 3. ИНИЦИАЛИЗАЦИЯ ---
data = {}
client_info = {}
score = 0

# --- 4. ШАПКА: ИНФОРМАЦИЯ О КЛИЕНТЕ ---
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

# --- 5. ТЕХНИЧЕСКИЙ ОПРОСНИК ---
st.header("📋 Технический аудит")

with st.expander("Блок 1: Инфраструктура", expanded=True):
    total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)
    data['АРМ'] = total_arm
    
    servers = st.number_input("Количество серверов (физических и виртуальных):", min_value=0, step=1)
    data['Серверы'] = servers
    
    mon = st.selectbox("Система мониторинга:", ["Нет", "Open-source", "Коммерческое ПО"])
    data['Мониторинг'] = mon

with st.expander("Блок 2: Информационная Безопасность"):
    ib_metrics = {
        "Резервное копирование": 20,
        "DLP (Защита от утечек)": 15,
        "PAM (Контроль доступа)": 10,
        "WAF (Защита Web)": 10,
        "EDR/Antimalware": 15
    }
    for item, pts in ib_metrics.items():
        if st.checkbox(item):
            vendor = st.text_input(f"Вендор {item}:", key=f"v_{item}")
            data[item] = f"Да ({vendor if vendor else 'не указан'})"
            score += pts
        else:
            data[item] = "Нет"

# --- 6. ЭКСПЕРТНАЯ ЛОГИКА ---
def get_recommendation(key, value):
    db = {
        "Нет": "Критический риск. Отсутствие данной системы снижает устойчивость бизнеса.",
        "Резервное копирование": "КРИТИЧНО: Рекомендуется стратегия 3-2-1. Без бэкапов данные под угрозой.",
        "Мониторинг": "Рекомендуется внедрение для сокращения времени простоя сервисов.",
        "EDR/Antimalware": "Рекомендуется защита класса EDR для борьбы с современными угрозами.",
        "DLP (Защита от утечек)": "Рекомендуется для контроля перемещения конфиденциальных данных."
    }
    if "Нет" in str(value) or value == 0:
        return db.get(key, db["Нет"])
    return "Конфигурация в норме. Рекомендуется регулярная актуализация."

# --- 7. ГЕНЕРАЦИЯ EXCEL ---
def make_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "IT Audit 2026"
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Заголовок
    ws.merge_cells('A1:D1')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ: АУДИТ ИТ И ИБ 2026"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center')

    # Шапка компании
    ws['A3'] = "КОМПАНИЯ:"; ws['B3'] = c_info['Наименование компании']
    ws['A4'] = "ИНДЕКС ЗРЕЛОСТИ:"; ws['B4'] = f"{final_score}/100"
    
    # Таблица результатов
    headers = ["Параметр", "Значение", "Анализ", "Рекомендация"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=6, column=i, value=h)
        cell.fill = header_fill; cell.font = white_font

    row = 7
    for k, v in results.items():
        ws.cell(row=row, column=1, value=k).border = border
        ws.cell(row=row, column=2, value=str(v)).border = border
        rec = get_recommendation(k, v)
        ws.cell(row=row, column=3, value="РИСК" if "Нет" in str(v) else "ОК").border = border
        ws.cell(row=row, column=4, value=rec).border = border
        ws.cell(row=row, column=4).alignment = Alignment(wrap_text=True)
        row += 1

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['D'].width = 50
    
    wb.save(output)
    return output.getvalue()

# --- 8. СИНХРОНИЗАЦИЯ С ОБЛАКОМ ---
def sync_to_google(c_info, results, final_score, excel_bytes):
    try:
        # Авторизация через Secrets
        creds_info = st.secrets["gcp_service_account"]
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_info, scopes=scope)
        
        # 1. Запись в Таблицу (CRM)
        gc = gspread.authorize(creds)
        sheet = gc.open(SHEET_NAME).sheet1
        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            c_info['Наименование компании'],
            c_info['Контактное лицо'],
            c_info['Контактный телефон'],
            f"{final_score}/100"
        ]
        sheet.append_row(row)
        
        # 2. Загрузка Excel в Папку
        drive = build('drive', 'v3', credentials=creds)
        filename = f"Audit_{c_info['Наименование компании']}_{datetime.now().strftime('%d%m%Y')}.xlsx"
        metadata = {'name': filename, 'parents': [FOLDER_ID]}
        media = MediaIoBaseUpload(BytesIO(excel_bytes), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        drive.files().create(body=metadata, media_body=media, fields='id').execute()
        
        return True
    except Exception as e:
        return str(e)

# --- 9. ФИНАЛЬНОЕ ДЕЙСТВИЕ ---
st.divider()
if st.button("🚀 Завершить аудит и отправить отчет"):
    if not client_info['Наименование компании']:
        st.error("Пожалуйста, заполните название компании!")
    else:
        f_score = min(score, 100)
        excel_data = make_excel(client_info, data, f_score)
        
        with st.spinner('Синхронизация данных с облаком...'):
            cloud_res = sync_to_google(client_info, data, f_score, excel_data)
            
            if cloud_res is True:
                st.success("✅ Данные сохранены в CRM и копия загружена на Google Диск!")
            else:
                st.error(f"Ошибка облака: {cloud_res}")
        
        st.download_button("📥 Скачать персональный отчет (Excel)", excel_data, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Разработано Ivan Rudoy. Данные передаются по защищенному каналу.")