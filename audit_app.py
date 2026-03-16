import streamlit as st
import pandas as pd
import requests
import os
from io import BytesIO
from datetime import datetime

# Библиотеки для Excel
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- 1. НАСТРОЙКИ ---
st.set_page_config(page_title="Аудит ИТ 2026", layout="wide", page_icon="🛡️")

# Данные Telegram из Secrets
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ШАПКА ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=250)
else:
    st.title("Khalil Trade | IT Audit")

st.markdown("### Экспертный анализ инфраструктуры")
st.divider()

# --- 3. ДАННЫЕ КЛИЕНТА ---
st.header("🏢 Информация о компании")
c1, c2 = st.columns(2)
with c1:
    company_name = st.text_input("Наименование компании")
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

# --- 4. ТЕХНИЧЕСКИЙ АУДИТ ---
st.header("📋 Технический аудит")
data = {}
score = 0

with st.expander("Инфраструктура и Безопасность", expanded=True):
    data['АРМ'] = st.number_input("Кол-во АРМ (шт):", min_value=0, step=1)
    data['Серверы'] = st.number_input("Кол-во серверов:", min_value=0, step=1)
    
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
            v_name = st.text_input(f"Вендор {task}:", key=f"v_{task}")
            data[task] = f"Да ({v_name if v_name else 'не указан'})"
            score += pts
        else:
            data[task] = "Нет"

# --- 5. ГЕНЕРАЦИЯ EXCEL ---
def make_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Report"
    
    # Оформление
    h_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    f_white = Font(color="FFFFFF", bold=True)
    side = Side(style='thin')
    brd = Border(left=side, right=side, top=side, bottom=side)

    # Заполнение
    ws.merge_cells('A1:D1')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ 2026"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center')

    ws['A3'] = "КОМПАНИЯ:"; ws['B3'] = c_info['Компания']
    ws['A4'] = "ИНДЕКС ЗРЕЛОСТИ:"; ws['B4'] = f"{final_score}/100"
    
    headers = ["Параметр", "Значение", "Статус", "Рекомендация"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=6, column=i, value=h)
        cell.fill = h_fill; cell.font = f_white; cell.border = brd

    row_num = 7
    for k, v in results.items():
        ws.cell(row=row_num, column=1, value=k).border = brd
        ws.cell(row=row_num, column=2, value=str(v)).border = brd
        status = "РИСК" if "Нет" in str(v) or v == 0 else "ОК"
        ws.cell(row=row_num, column=3, value=status).border = brd
        ws.cell(row=row_num, column=4, value="Требуется аудит" if status == "РИСК" else "В норме").border = brd
        row_num += 1

    ws.column_dimensions['A'].width = 25; ws.column_dimensions['D'].width = 40
    wb.save(output)
    return output.getvalue()

# --- 6. ОТПРАВКА В TELEGRAM ---
def send_telegram(c_info, final_score, excel_bytes):
    try:
        msg = (f"🔔 *НОВЫЙ АУДИТ*\n\n🏢 *Компания:* {c_info['Компания']}\n"
               f"👤 *Контакт:* {c_info['Лицо']}\n📊 *Балл:* {final_score}/100")
        
        # Текст
        requests.post(f"https://api.telegram.org/bot{TOKEN}/sendMessage", 
                      data={"chat_id": CHAT_ID, "text": msg, "parse_mode": "Markdown"})
        
        # Файл
        requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                      data={"chat_id": CHAT_ID}, 
                      files={'document': (f"Audit_{c_info['Компания']}.xlsx", excel_bytes)})
        return True
    except Exception as e:
        return str(e)

# --- 7. КНОПКА ЗАПУСКА ---
st.divider()
if st.button("🚀 Сформировать и отправить отчет"):
    if not company_name:
        st.error("Введите название компании!")
    elif not TOKEN:
        st.error("Настройте TELEGRAM_TOKEN в Secrets!")
    else:
        final_score = min(score, 100)
        excel_data = make_excel(client_info, data, final_score)
        
        with st.spinner('Отправка в Telegram...'):
            res = send_telegram(client_info, final_score, excel_data)
            if res is True:
                st.success("✅ Отчет отправлен в Telegram!")
                st.download_button("📥 Скачать Excel", excel_data, f"Audit_{company_name}.xlsx")
            else:
                st.error(f"Ошибка: {res}")
