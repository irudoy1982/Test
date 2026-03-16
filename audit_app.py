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

# --- КОНФИГУРАЦИЯ ---
SHEET_NAME = "Audit_Results_2026"
FOLDER_ID = "ВАШ_ID_ПАПКИ" # Замените на ID вашей папки из Google Drive

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

# --- 2. ДИАГНОСТИКА СЕТИ ---
def check_network():
    st.sidebar.subheader("Статус системы")
    try:
        # Проверяем доступность серверов Google
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

st.markdown("### Экспертный аудит инфраструктуры")
st.divider()

# --- 4. ДАННЫЕ КЛИЕНТА ---
st.header("🏢 Общая информация")
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
    'Наименование компании': company_name,
    'Сайт Компании': site,
    'Контактное лицо': contact_person,
    'Должность': position,
    'Контактный email': email,
    'Контактный телефон': phone
}

st.divider()

# --- 5. ТЕХНИЧЕСКИЙ ОПРОСНИК ---
st.header("📋 Технический аудит")
data = {}
score = 0

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
            vendor = st.text_input(f"Вендор {task}:", key=f"v_{task}")
            data[task] = f"Да ({vendor if vendor else 'не указан'})"
            score += pts
        else:
            data[task] = "Нет"

# Вспомогательная функция для рекомендаций
def get_recommendation(key, value):
    if "Нет" in str(value) or value == 0:
        return "Критический риск. Требуется внедрение системы для обеспечения безопасности."
    return "В норме. Рекомендуется плановая проверка настроек."

# --- 6. ГЕНЕРАЦИЯ EXCEL ---
def make_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Report"
    
    # Стили оформления
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Заголовок отчета
    ws.merge_cells('A1:D1')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ 2026"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center')

    # Информация о клиенте
    ws['A3'] = "КОМПАНИЯ:"; ws['B3'] = c_info['Наименование компании']
    ws['A4'] = "ИНДЕКС ЗРЕЛОСТИ:"; ws['B4'] = f"{final_score}/100"
    
    # Заголовки таблицы
    headers = ["Параметр", "Значение", "Статус", "Рекомендация"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=6, column=i, value=h)
        cell.fill = header_fill
        cell.font = white_font
        cell.border = border

    # Заполнение данными
    curr_row = 7
    for k, v in results.items():
        ws.cell(row=curr_row, column=1, value=k).border = border
        ws.cell(row=curr_row, column=2, value=str(v)).border = border
        
        status = "РИСК" if "Нет" in str(v) or v == 0 else "ОК"
        st_cell = ws.cell(row=curr_row, column=3, value=status)
        st_cell.border = border
        if status == "РИСК":
            st_cell.font = Font(color="FF0000", bold=True)
            
        ws.cell(row=curr_row, column=4, value=get_recommendation(k, v)).border = border
        ws.cell(row=curr_row, column=4).alignment = Alignment(wrap_text=True)
        curr_row += 1

    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['D'].width = 50
    
    wb.save(output)
    return output.getvalue()

# --- 7. ОБЛАЧНАЯ СИНХРОНИЗАЦИЯ ---
def sync_to_google(c_info, results, final_score, excel_bytes):
    try:
        # 1. Проверка наличия секретов
        if "gcp_service_account" not in st.secrets:
            return "Секреты gcp_service_account не найдены!"

        # 2. Подготовка словаря ключей
        creds_dict = dict(st.secrets["gcp_service_account"])
        # Исправляем формат приватного ключа
        if "\\n" in creds_dict["private_key"]:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")

        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
        
        # 3. Запись в Таблицу
        gc = gspread.authorize(creds)
        sheet = gc.open(SHEET_NAME).sheet1
        row_data = [
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            c_info['Наименование компании'],
            c_info['Контактное лицо'],
            c_info['Контактный телефон'],
            f"{final_score}/100"
        ]
        sheet.append_row(row_data)
        
        # 4. Загрузка Excel в Папку
        drive_service = build('drive', 'v3', credentials=creds)
        fname = f"Audit_{c_info['Наименование компании']}_{datetime.now().strftime('%d%m%Y')}.xlsx"
        metadata = {'name': fname, 'parents': [FOLDER_ID]}
        media = MediaIoBaseUpload(BytesIO(excel_bytes), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        drive_service.files().create(body=metadata, media_body=media, fields='id').execute()
        
        return True
    except Exception as e:
        return str(e)

# --- 8. ФИНАЛЬНАЯ КНОПКА ---
st.divider()
if st.button("🚀 Сформировать и отправить отчет в облако"):
    if not is_online:
        st.error("Ошибка сети: Нет связи с Google. Проверьте подключение или сделайте Reboot App.")
    elif not company_name:
        st.error("Введите название компании!")
    else:
        final_score = min(score, 100)
        excel_data = make_excel(client_info, data, final_score)
        
        with st.spinner('Синхронизация с облаком Google...'):
            result = sync_to_google(client_info, data, final_score, excel_data)
            
            if result is True:
                st.success("✅ Отчет успешно сохранен в CRM (Таблицу) и загружен на Диск!")
            else:
                st.error(f"Ошибка при сохранении: {result}")
        
        st.download_button(
            label="📥 Скачать копию Excel",
            data=excel_data,
            file_name=f"Audit_{company_name}.xlsx"
        )

st.info("Разработка: Ivan Rudoy. Система готова к работе.")
