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

# --- 2. ЛОГОТИП И ШАПКА ---
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

# --- 3. ОБЩАЯ ИНФОРМАЦИЯ (ШАПКА) ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город:*")
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    
    # Логика работы с сайтом и доменом
    site_input = st.text_input("Сайт компании (например, khalil.kz):", key="site_field")
    client_info['Сайт компании'] = site_input

    # Извлекаем чистый домен для Email
    clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]

    if clean_domain:
        st.write("**Email контактного лица:**")
        e_col1, e_col2 = st.columns([3, 2])
        with e_col1:
            email_prefix = st.text_input("Введите логин (до @):", placeholder="info", label_visibility="collapsed", key="email_pre")
        with e_col2:
            st.markdown(f"<div style='padding-top: 5px; font-size: 18px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
        client_info['Email'] = f"{email_prefix}@{clean_domain}" if email_prefix else ""
    else:
        st.info("ℹ️ Введите сайт компании, чтобы зафиксировать домен почты")
        client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    client_info['Контактный телефон'] = st.text_input("Контактный телефон:*")

st.divider()

# --- 4. ТЕХНИЧЕСКИЕ БЛОКИ ---
# Блок 1: Инфраструктура
st.header("Блок 1: Техническая часть")
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)
data['1.1. Всего АРМ'] = total_arm

st.subheader("1.4. Почтовая система")
data['1.4. Почта'] = st.selectbox("Тип почты:", ["Microsoft 365", "Google Workspace", "Exchange (On-Prem)", "SmarterMail", "Собственный Linux-сервер", "Нет"])

# Блок 3: ИБ
st.header("Блок 3: Информационная Безопасность")
ib_items = {
    "DLP (Защита от утечек)": 15,
    "PAM (Контроль доступа)": 10,
    "SIEM/SOC (Мониторинг ИБ)": 20,
    "WAF (Защита Web)": 10,
    "EDR/Antimalware": 15,
    "Резервное копирование": 20
}
for label, pts in ib_items.items():
    c1, c2 = st.columns([1, 2])
    if c1.checkbox(label):
        v_name = c2.text_input(f"Вендор {label}:", key=f"v_{label}")
        data[label] = f"Да ({v_name if v_name else 'не указан'})"
        score += pts
    else:
        data[label] = "Нет"

# Блок 4: Web
st.header("Блок 4: Web-ресурсы")
if st.toggle("Есть свои Web-ресурсы"):
    data['4.1. Хостинг'] = st.selectbox("Где размещены сайты:", ["Собственный ЦОД", "Облако (KZ)", "Облако (Global)"])
    data['4.4. Frontend Web-серверы'] = st.multiselect(
        "Какие веб-сервера используются на frontend?", 
        ["Nginx", "Apache HTTP Server", "Microsoft IIS", "LiteSpeed", "Caddy", "Cloudflare", "Другое"]
    )

# --- 5. ГЕНЕРАЦИЯ EXCEL ---
def make_expert_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Khalil Audit Report"

    # Стили
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Шапка отчета
    ws.merge_cells('A1:D2')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ: ТЕХНИЧЕСКИЙ АУДИТ 2026"
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=14, color="1F4E78")

    # Данные клиента
    current_row = 4
    for k, v in c_info.items():
        ws.cell(row=current_row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=current_row, column=2, value=str(v))
        current_row += 1
    
    # Автоматическая дата
    report_date = datetime.now().strftime("%d.%m.%Y %H:%M")
    ws.cell(row=current_row, column=1, value="Дата аудита (авто):").font = Font(bold=True)
    ws.cell(row=current_row, column=2, value=report_date)
    current_row += 2

    # Индекс зрелости
    ws.cell(row=current_row, column=1, value="ИНДЕКС ТЕХНИЧЕСКОЙ ЗРЕЛОСТИ:").font = Font(bold=True)
    score_cell = ws.cell(row=current_row, column=2, value=f"{final_score}%")
    bg_color = "92D050" if final_score > 70 else "FFC000" if final_score > 40 else "FF7C80"
    score_cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    score_cell.font = Font(bold=True)
    current_row += 2

    # Таблица параметров
    headers = ["Параметр", "Значение", "Статус", "Рекомендация эксперта"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=current_row, column=i, value=h)
        cell.fill = header_fill; cell.font = white_font

    current_row += 1
    rec_map = {
        "Нет": "Требуется внедрение для минимизации бизнес-рисков.",
        "Резервное копирование": "Настроить автоматическое копирование по правилу 3-2-1.",
        "SIEM/SOC": "Рекомендуется для раннего обнаружения кибератак.",
        "Email": "Обеспечить защиту корпоративной почты (SPF, DKIM, DMARC)."
    }

    for k, v in results.items():
        ws.cell(row=current_row, column=1, value=k).border = border
        ws.cell(row=current_row, column=2, value=str(v)).border = border
        
        is_risk = "Нет" in str(v) or v == 0 or v == []
        status_text = "РИСК" if is_risk else "ОК"
        st_cell = ws.cell(row=current_row, column=3, value=status_text)
        st_cell.border = border
        if is_risk: st_cell.font = Font(color="FF0000", bold=True)

        rec_text = rec_map.get(k, "Поддерживать текущую работоспособность.") if not is_risk else rec_map.get(k, "Рассмотреть внедрение решения.")
        ws.cell(row=current_row, column=4, value=rec_text).border = border
        current_row += 1

    # Логотип (если есть)
    if os.path.exists("logo.png"):
        try:
            img = OpenpyxlImage("logo.png")
            img.height = 55; img.width = 160
            ws.add_image(img, 'D1')
        except: pass

    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['D'].width = 50
    
    wb.save(output)
    return output.getvalue(), report_date

# --- 6. КНОПКА ОТПРАВКИ ---
st.divider()
if st.button("📊 Сформировать экспертный отчет", use_container_width=True):
    mandatory = [client_info['Город'], client_info['Наименование компании'], client_info['ФИО контактного лица']]
    if not all(mandatory):
        st.error("⚠️ Заполните все обязательные поля (отмечены *)!")
    elif not client_info['Email']:
        st.error("⚠️ Не заполнен Email или не указан сайт компании!")
    else:
        with st.spinner("Создаем отчет и уведомляем команду..."):
            f_score = min(score, 100)
            report_bytes, final_date = make_expert_excel(client_info, data, f_score)
            
            try:
                # Текст для Telegram
                caption = (f"🚀 *Коллеги, у нас новый заказ. Давайте зарабатывать!*\n\n"
                           f"🏢 *Компания:* {client_info['Наименование компании']}\n"
                           f"📧 *Email:* {client_info['Email']}\n"
                           f"📊 *Зрелость ИТ:* {f_score}%\n"
                           f"📅 *Дата:* {final_date}\n"
                           f"👤 *Контакт:* {client_info['ФИО контактного лица']}")
                
                url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
                files = {'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report_bytes)}
                requests.post(url, data={"chat_id": CHAT_ID, "caption": caption, "parse_mode": "Markdown"}, files=files)
                
                st.success("Отчет успешно отправлен в Telegram!")
                st.balloons()
            except Exception as e:
                st.error(f"Ошибка при отправке в Telegram: {e}")
            
            st.download_button("📥 Скачать готовый Excel-отчет", report_bytes, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v2.1 | Almaty 2026")
