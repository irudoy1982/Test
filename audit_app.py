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

# --- НАСТРОЙКИ TELEGRAM (из Secrets) ---
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ЛОГОТИП И КОНТАКТЫ ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### Мы поможем Вам стать лучше!**")
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
    
    # ПОЛЕ САЙТ
    site_input = st.text_input("Сайт компании (например, khalil.kz):", key="site_field")
    client_info['Сайт компании'] = site_input

    # ПОЛЕ EMAIL С АВТОПОДСТАНОВКОЙ
    domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
    suffix = f"@{domain}" if domain else ""
    
    email_prefix = st.text_input(f"Email контактного лица (префикс {suffix}):", key="email_field")
    client_info['Email'] = f"{email_prefix}{suffix}" if email_prefix else ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    client_info['Контактный телефон'] = st.text_input("Контактный телефон:*")

st.divider()

# --- ТЕХНИЧЕСКИЕ БЛОКИ (Кратко для экономии места) ---
st.header("Блок 1: Техническая часть")
data['1.1. Всего АРМ'] = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)
data['1.4. Почта'] = st.selectbox("Тип почты:", ["Microsoft 365", "Google Workspace", "SmarterMail", "Exchange", "Нет"])

st.header("Блок 3: Информационная Безопасность")
ib_list = {"DLP": 15, "PAM": 10, "SIEM": 20, "Резервное копирование": 20}
for label, pts in ib_list.items():
    if st.checkbox(label):
        v = st.text_input(f"Вендор {label}:")
        data[label] = f"Да ({v})"
        score += pts
    else:
        data[label] = "Нет"

st.header("Блок 4: Web-ресурсы")
if st.toggle("Есть свои Web-ресурсы"):
    data['4.4. Frontend'] = st.multiselect("Frontend серверы:", ["Nginx", "Apache", "IIS", "Cloudflare"])

# --- ГЕНЕРАЦИЯ EXCEL ---
def make_expert_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Khalil Audit Report"
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Заголовок
    ws.merge_cells('A1:D2')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ ПО ИТ И ИБ (2026)"
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    
    # Инфо о клиенте
    current_row = 4
    for k, v in c_info.items():
        ws.cell(row=current_row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=current_row, column=2, value=str(v))
        current_row += 1
    
    auto_date = datetime.now().strftime("%d.%m.%Y %H:%M")
    ws.cell(row=current_row, column=1, value="Дата генерации:").font = Font(bold=True)
    ws.cell(row=current_row, column=2, value=auto_date)
    current_row += 2

    # Скоринг
    ws.cell(row=current_row, column=1, value="ИНДЕКС ТЕХНИЧЕСКОЙ ЗРЕЛОСТИ:").font = Font(bold=True)
    score_cell = ws.cell(row=current_row, column=2, value=f"{final_score}%")
    bg_color = "92D050" if final_score > 70 else "FFC000" if final_score > 40 else "FF7C80"
    score_cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    current_row += 2

    # Таблица результатов
    headers = ["Параметр", "Значение", "Статус", "Рекомендация"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=current_row, column=i, value=h)
        cell.fill = header_fill; cell.font = white_font

    current_row += 1
    for k, v in results.items():
        ws.cell(row=current_row, column=1, value=k).border = border
        ws.cell(row=current_row, column=2, value=str(v)).border = border
        ws.cell(row=current_row, column=3, value="ОК" if "Да" in str(v) else "РИСК").border = border
        ws.cell(row=current_row, column=4, value="Рекомендовано Khalil Trade").border = border
        current_row += 1

    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 30
    wb.save(output)
    return output.getvalue(), auto_date

# --- ОТПРАВКА ---
st.divider()
if st.button("📊 Сформировать экспертный отчет"):
    if not client_info['Email'] or "@" not in client_info['Email']:
        st.error("⚠️ Пожалуйста, корректно заполните сайт и префикс Email!")
    else:
        with st.spinner("Отправка данных..."):
            f_score = min(score, 100)
            report_bytes, final_date = make_expert_excel(client_info, data, f_score)
            
            try:
                url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
                caption = (f"🚀 *Коллеги, у нас новый заказ. Давайте зарабатывать!*\n\n"
                           f"🏢 *Компания:* {client_info['Наименование компании']}\n"
                           f"📧 *Email:* {client_info['Email']}\n"
                           f"📊 *Зрелость:* {f_score}%\n"
                           f"📅 *Дата:* {final_date}")
                
                requests.post(url, data={"chat_id": CHAT_ID, "caption": caption, "parse_mode": "Markdown"}, 
                              files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report_bytes)})
                st.success(f"Отчет для {client_info['Email']} отправлен!")
            except Exception as e:
                st.error(f"Ошибка: {e}")
            
            st.download_button("📥 Скачать Excel", report_bytes, "Audit_Khalil_2026.xlsx")

st.info("Khalil Audit | Almaty 2026")
