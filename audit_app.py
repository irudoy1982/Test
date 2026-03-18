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

# --- НАСТРОЙКИ TELEGRAM ---
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ЛОГОТИП ---
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

# --- 3. ОБЩАЯ ИНФОРМАЦИЯ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город:*")
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    site_input = st.text_input("Сайт компании:", key="site_field")
    client_info['Сайт компании'] = site_input

    # Логика Email с жестким доменом
    clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
    if clean_domain:
        st.write("Email контактного лица:*")
        e_col1, e_col2 = st.columns([2, 3])
        with e_col1:
            email_prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="email_pre")
        with e_col2:
            st.markdown(f"<div style='padding-top: 5px; font-size: 16px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
        client_info['Email'] = f"{email_prefix}@{clean_domain}" if email_prefix else ""
    else:
        client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    
    # Компактное поле телефона
    st.write("Контактный телефон:*")
    t_col1, t_col2 = st.columns([1, 12])
    with t_col1:
        st.markdown("<div style='padding-top: 5px; font-size: 18px; font-weight: bold;'>+7</div>", unsafe_allow_html=True)
    with t_col2:
        phone_main = st.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed", key="phone_input")
    client_info['Контактный телефон'] = f"+7 {phone_main}" if phone_main else ""

st.divider()

# --- 4. ТЕХНИЧЕСКИЕ БЛОКИ ---

# БЛОК 1: АРМ
st.header("Блок 1: Конечные точки (АРМ)")
data['1.1. Всего АРМ'] = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)
data['1.4. Почта'] = st.selectbox("Тип почты:", ["Microsoft 365", "Google Workspace", "Exchange", "SmarterMail", "Нет"])

# БЛОК 2: СЕТЬ (ВТОРОЙ)
st.header("Блок 2: Сетевая инфраструктура")
if st.toggle("Показать настройки сети", key="net_tgl"):
    net_opts = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink"]
    data['2.1. Основной канал'] = st.selectbox("Тип канала:", net_opts)
    data['2.4. NGFW'] = st.text_input("Вендор Межсетевого экрана (NGFW):")
    if data['2.4. NGFW']: score += 20

# БЛОК 3: СЕРВЕРЫ (ТРЕТИЙ)
st.header("Блок 3: Серверная инфраструктура")
if st.toggle("Показать настройки серверов", key="srv_tgl"):
    cs1, cs2 = st.columns(2)
    data['3.1. Физические серверы'] = cs1.number_input("Физических серверов:", min_value=0, step=1)
    data['3.2. Виртуальные серверы'] = cs2.number_input("Виртуальных серверов:", min_value=0, step=1)
    data['3.3. Системы виртуализации'] = st.multiselect("Используются:", ["VMware", "Hyper-V", "Proxmox", "KVM", "Нет"])

# БЛОК 4: ИБ
st.header("Блок 4: Информационная Безопасность")
ib_list = {"DLP": 15, "PAM": 10, "SIEM/SOC": 20, "Резервное копирование": 20}
for label, pts in ib_list.items():
    c1, c2 = st.columns([1, 2])
    if c1.checkbox(label, key=f"chk_{label}"):
        v_name = c2.text_input(f"Вендор {label}:", key=f"vn_{label}")
        data[label] = f"Да ({v_name if v_name else 'не указан'})"
        score += pts
    else:
        data[label] = "Нет"

# --- 5. ГЕНЕРАЦИЯ EXCEL ---
def make_expert_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Khalil Audit Report"

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Шапка
    ws.merge_cells('A1:D2')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ: ТЕХНИЧЕСКИЙ АУДИТ 2026"
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=14, color="1F4E78")

    # Инфо
    curr_row = 4
    for k, v in c_info.items():
        ws.cell(row=curr_row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=curr_row, column=2, value=str(v))
        curr_row += 1
    
    rep_date = datetime.now().strftime("%d.%m.%Y %H:%M")
    ws.cell(row=curr_row, column=1, value="Дата отчета:").font = Font(bold=True)
    ws.cell(row=curr_row, column=2, value=rep_date)
    curr_row += 2

    # Скоринг
    ws.cell(row=curr_row, column=1, value="ИНДЕКС ТЕХНИЧЕСКОЙ ЗРЕЛОСТИ:").font = Font(bold=True)
    score_cell = ws.cell(row=curr_row, column=2, value=f"{final_score}%")
    bg_color = "92D050" if final_score > 70 else "FFC000" if final_score > 40 else "FF7C80"
    score_cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    curr_row += 2

    # Таблица данных
    headers = ["Параметр", "Значение", "Статус", "Рекомендация эксперта"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=curr_row, column=i, value=h)
        cell.fill = header_fill; cell.font = Font(color="FFFFFF", bold=True)

    curr_row += 1
    for k, v in results.items():
        ws.cell(row=curr_row, column=1, value=k).border = border
        ws.cell(row=curr_row, column=2, value=str(v)).border = border
        is_risk = "Нет" in str(v) or v == 0
        ws.cell(row=curr_row, column=3, value="РИСК" if is_risk else "ОК").border = border
        ws.cell(row=curr_row, column=4, value="Рекомендовано Khalil Trade").border = border
        curr_row += 1

    if os.path.exists("logo.png"):
        try:
            img = OpenpyxlImage("logo.png")
            img.height = 55; img.width = 160
            ws.add_image(img, 'D1')
        except: pass

    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['D'].width = 40
    wb.save(output)
    return output.getvalue(), rep_date

# --- 6. КНОПКА ОТПРАВКИ ---
st.divider()
if st.button("📊 Сформировать экспертный отчет", use_container_width=True):
    mand = [client_info['Город'], client_info['Наименование компании'], client_info['ФИО контактного лица']]
    if not all(mand):
        st.error("⚠️ Заполните Город, Компанию и ФИО!")
    elif not client_info['Email'] or not phone_main:
        st.error("⚠️ Укажите Email и Телефон!")
    else:
        with st.spinner("Создаем отчет..."):
            f_score = min(score, 100)
            report_bytes, final_date = make_expert_excel(client_info, data, f_score)
            
            try:
                msg = (f"🚀 *Новый аудит!*\n\n"
                       f"🏢 *Компания:* {client_info['Наименование компании']}\n"
                       f"📧 *Email:* {client_info['Email']}\n"
                       f"📞 *Тел:* {client_info['Контактный телефон']}\n"
                       f"📊 *Зрелость:* {f_score}%\n"
                       f"📅 *Дата:* {final_date}")
                
                requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                              data={"chat_id": CHAT_ID, "caption": msg, "parse_mode": "Markdown"}, 
                              files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report_bytes)})
                st.success("Отчет отправлен!")
                st.balloons()
            except:
                st.error("Ошибка Telegram API")
            
            st.download_button("📥 Скачать Excel", report_bytes, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System | Almaty 2026")
