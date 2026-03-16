import streamlit as st
import pandas as pd
import requests
import os
from io import BytesIO
from datetime import datetime

# Библиотеки для Excel
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

# Получение настроек из Secrets (Streamlit Cloud)
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ЛОГОТИП И ШАПКА ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### Экспертная оценка ИТ-инфраструктуры")
st.divider()

# --- 3. ИНФОРМАЦИЯ О КЛИЕНТЕ ---
st.header("🏢 Общая информация")
c1, c2 = st.columns(2)
with c1:
    company_name = st.text_input("Наименование компании", placeholder="ТОО 'Пример'")
    contact_person = st.text_input("Контактное лицо (ФИО)")
with c2:
    email = st.text_input("Контактный email")
    phone = st.text_input("Контактный телефон")

client_info = {
    "Компания": company_name,
    "Лицо": contact_person,
    "Телефон": phone,
    "Email": email
}

st.divider()

# --- 4. ТЕХНИЧЕСКИЙ ОПРОСНИК ---
st.header("📋 Технический аудит")
data = {}
score = 0

with st.expander("Инфраструктура и Безопасность", expanded=True):
    data['АРМ'] = st.number_input("Количество АРМ (шт):", min_value=0, step=1)
    data['Серверы'] = st.number_input("Количество серверов:", min_value=0, step=1)
    
    # Системы ИБ и их вес в баллах
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
            vendor = st.text_input(f"Вендор {task}:", key=f"v_{task}")
            data[task] = f"Да ({vendor if vendor else 'не указан'})"
            score += pts
        else:
            data[task] = "Нет"

# --- 5. ГЕНЕРАЦИЯ EXCEL ---
def make_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Report"
    
    # Стили (Цвет шапки, шрифты, границы)
    h_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    f_white = Font(color="FFFFFF", bold=True)
    side = Side(style='thin')
    brd = Border(left=side, right=side, top=side, bottom=side)

    # Заголовок отчета
    ws.merge_cells('A1:D1')
    ws['A1'] = "ИТ-АУДИТ И КИБЕРБЕЗОПАСНОСТЬ 2026"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center')

    ws['A3'] = "КОМПАНИЯ:"; ws['B3'] = c_info['Компания']
    ws['A4'] = "ИНДЕКС ЗРЕЛОСТИ:"; ws['B4'] = f"{final_score}/100"
    
    # Заголовки таблицы
    headers = ["Параметр", "Значение", "Статус", "Рекомендация"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=6, column=i, value=h)
        cell.fill = h_fill
        cell.font = f_white
        cell.border = brd

    # Заполнение данными
    curr_row = 7
    for k, v in results.items():
        ws.cell(row=curr_row, column=1, value=k).border = brd
        ws.cell(row=curr_row, column=2, value=str(v)).border = brd
        
        is_risk = "Нет" in str(v) or v == 0
        status = "РИСК" if is_risk else "ОК"
        
        st_cell = ws.cell(row=curr_row, column=3, value=status)
        st_cell.border = brd
        if is_risk:
            st_cell.font = Font(color="FF0000", bold=True)
            
        rec = "Требуется внедрение системы" if is_risk else "Система внедрена"
        ws.cell(row=curr_row, column=4, value=rec).border = brd
        curr_row += 1

    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['D'].width = 30
    
    wb.save(output)
    return output.getvalue()

# --- 6. ОТПРАВКА В TELEGRAM С ДИАГНОСТИКОЙ ---
def send_telegram(c_info, final_score, excel_bytes):
    try:
        url_base = f"https://api.telegram.org/bot{TOKEN}"
        
        # 1. Текстовое сообщение
        msg = (f"🔔 *НОВЫЙ АУДИТ*\n\n"
               f"🏢 *Компания:* {c_info['Компания']}\n"
               f"👤 *Контакт:* {c_info['Лицо']}\n"
               f"📊 *Балл:* {final_score}/100")
        
        r_text = requests.post(f"{url_base}/sendMessage", 
                               data={"chat_id": CHAT_ID, "text": msg, "parse_mode": "Markdown"})
        
        # 2. Отправка файла Excel
        files = {'document': (f"Audit_{c_info['Компания']}.xlsx", excel_bytes)}
        r_file = requests.post(f"{url_base}/sendDocument", 
                               data={"chat_id": CHAT_ID}, files=files)
        
        # Проверка результата
        if r_text.ok and r_file.ok:
            return True
        else:
            # Если Telegram вернул ошибку, вытягиваем её описание
            err_desc = r_text.json().get('description', 'Неизвестная ошибка API')
            return f"Telegram отклонил запрос: {err_desc}"
            
    except Exception as e:
        return f"Ошибка соединения: {str(e)}"

# --- 7. ФИНАЛЬНОЕ ДЕЙСТВИЕ ---
st.divider()
if st.button("🚀 Сформировать и отправить отчет"):
    if not company_name:
        st.error("Пожалуйста, введите название компании!")
    elif not TOKEN or not CHAT_ID:
        st.error("Настройки Telegram (Token или Chat ID) не найдены в Secrets!")
    else:
        f_score = min(score, 100)
        excel_data = make_excel(client_info, data, f_score)
        
        with st.spinner('Связь с Telegram...'):
            res = send_telegram(client_info, f_score, excel_data)
            
            if res is True:
                st.success("✅ Отчет успешно отправлен в ваш Telegram!")
                st.download_button("📥 Скачать копию Excel", excel_data, f"Audit_{company_name}.xlsx")
