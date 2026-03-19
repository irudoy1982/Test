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

# --- НАСТРОЙКИ TELEGRAM (подтягиваются из .streamlit/secrets.toml) ---
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ВИЗУАЛ: ЛОГОТИП И ЗАГОЛОВОК ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### Мы поможем Вам стать лучше!")
st.divider()

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026)")

# Инициализация переменных для сбора данных
data = {}
client_info = {}
score = 0

# --- БЛОК: ОБЩАЯ ИНФОРМАЦИЯ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город:*")
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    
    # Логика Сайта
    site_input = st.text_input("Сайт компании:*", placeholder="example.kz", key="site_f")
    client_info['Сайт компании'] = site_input

    # Логика Email с автоматическим подставлением домена
    custom_email_mode = st.checkbox("Email отличается от домена сайта")
    
    if custom_email_mode:
        client_info['Email'] = st.text_input("Email контактного лица:*", placeholder="ivan@khalil.kz")
    else:
        # Извлекаем чистый домен для красоты
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица (логин):*")
            e_col1, e_col2 = st.columns([2, 3])
            with e_col1:
                email_prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="em_prefix")
            with e_col2:
                st.markdown(f"<div style='padding-top: 5px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
            client_info['Email'] = f"{email_prefix}@{clean_domain}" if email_prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    
    # --- НОВЫЙ БЛОК ТЕЛЕФОНА (Версия 2.0) ---
    st.write("Контактный телефон:*")
    p_col1, p_col2 = st.columns([1, 2])
    with p_col1:
        country_code = st.selectbox(
            "Код",
            options=[
                ("🇰🇿 +7", "+7"), 
                ("🇷🇺 +7", "+7"), 
                ("🇺🇿 +998", "+998"), 
                ("🇰🇬 +996", "+996"), 
                ("🇦🇪 +971", "+971")
            ],
            format_func=lambda x: x[0],
            label_visibility="collapsed"
        )
    with p_col2:
        phone_raw = st.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed", key="phone_val")
    
    client_info['Контактный телефон'] = f"{country_code[1]} {phone_raw}"

st.divider()

# --- БЛОК 1: ИТ ИНФРАСТРУКТУРА ---
st.header("Блок 1: Информационные технологии")

# 1.1 АРМ
st.subheader("1.1. Конечные точки (АРМ)")
data['1.1. Всего АРМ'] = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)
selected_os = st.multiselect("ОС на АРМ:", ["Windows", "Linux", "macOS", "Другое"])
for os_type in selected_os:
    data[f"Кол-во {os_type}"] = st.number_input(f"Сколько устройств на {os_type}?", min_value=0, step=1)

# 1.2 Сеть
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_t"):
    data['1.2.1. Канал'] = st.selectbox("Тип канала:", ["Оптика", "Радиоканал", "Спутник", "4G/5G", "Starlink"])
    data['1.2.2. NGFW'] = st.text_input("Вендор Межсетевого экрана (NGFW):", placeholder="Fortinet, Cisco, и т.д.")
    if data['1.2.2. NGFW']: score += 20
else:
    data['1.2. Сетевая инфраструктура'] = "Аренда/Нет"

# 1.3 Серверы и сервисы
st.write("---")
col_s1, col_s2 = st.columns(2)
with col_s1:
    data['1.3.1. Физические серверы'] = st.number_input("Физические серверы (шт):", min_value=0, step=1)
    data['1.4. Виртуализация'] = st.multiselect("Системы виртуализации:", ["VMware", "Hyper-V", "Proxmox", "KVM", "Нет"])
with col_s2:
    data['1.3.2. Виртуальные серверы'] = st.number_input("Виртуальные серверы (шт):", min_value=0, step=1)
    data['1.5. Почта'] = st.selectbox("Почтовая система:", ["Exchange", "M365", "Google", "Yandex", "Свой сервер", "Нет"])

data['1.6. Мониторинг'] = st.selectbox("Система мониторинга:", ["Нет", "Zabbix", "Nagios", "PRTG", "Prometheus"])

st.divider()

# --- БЛОК 2: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ ---
st.header("Блок 2: Информационная Безопасность")
if st.toggle("Используются специализированные системы ИБ", key="ib_t"):
    ib_map = {"DLP": 15, "PAM": 10, "SIEM": 20, "EDR": 15, "Backup": 20}
    for key, pts in ib_map.items():
        if st.checkbox(f"Система {key}", key=f"chk_{key}"):
            vendor = st.text_input(f"Вендор {key}:", key=f"v_{key}")
            data[key] = f"Да ({vendor if vendor else 'не указан'})"
            score += pts
        else:
            data[key] = "Нет"

# --- ФУНКЦИЯ СОЗДАНИЯ EXCEL ---
def create_report(c_info, results, total_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Khalil Audit"
    
    # Оформление
    blue_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    ws.merge_cells('A1:D2')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ Khalil Trade (2026)"
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=14, color="1F4E78")

    # Вставка лого если есть
    if os.path.exists("logo.png"):
        try:
            img = OpenpyxlImage("logo.png")
            img.height = 50; img.width = 150
            ws.add_image(img, 'D1')
        except: pass

    row = 4
    for k, v in c_info.items():
        ws.cell(row=row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=row, column=2, value=str(v))
        row += 1
    
    row += 1
    ws.cell(row=row, column=1, value="ИТ-ЗРЕЛОСТЬ:").font = Font(bold=True)
    ws.cell(row=row, column=2, value=f"{total_score}%").font = Font(bold=True)
    
    row += 2
    headers = ["Параметр", "Значение", "Статус", "Рекомендация"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=i, value=h)
        cell.fill = blue_fill; cell.font = white_font
    
    row += 1
    for k, v in results.items():
        ws.cell(row=row, column=1, value=k).border = border
        ws.cell(row=row, column=2, value=str(v)).border = border
        ws.cell(row=row, column=3, value="Проверено").border = border
        ws.cell(row=row, column=4, value="Оптимизировать по регламенту").border = border
        row += 1

    for col, w in {'A': 30, 'B': 30, 'C': 15, 'D': 40}.items():
        ws.column_dimensions[col].width = w
        
    wb.save(output)
    return output.getvalue()

# --- КНОПКА ОТПРАВКИ ---
st.divider()
if st.button("📊 Сформировать и отправить отчет", type="primary"):
    # Проверка на заполнение полей
    is_phone_ok = len(phone_raw.strip()) > 5
    is_mail_ok = "@" in client_info.get('Email', "")
    
    if not all([client_info['Наименование компании'], client_info['ФИО контактного лица']]) or not is_phone_ok or not is_mail_ok:
        st.error("⚠️ Заполните все обязательные поля (Компания, ФИО, Телефон и корректный Email)!")
    else:
        with st.spinner("Генерируем отчет и отправляем в Telegram..."):
            final_score = min(score, 100)
            excel_file = create_report(client_info, data, final_score)
            
            try:
                # Текст для Telegram
                text = (f"🚀 *Новый аудит: {client_info['Наименование компании']}*\n\n"
                        f"👤 *Контакт:* {client_info['ФИО контактного лица']}\n"
                        f"📞 *Телефон:* {client_info['Контактный телефон']}\n"
                        f"📧 *Email:* {client_info['Email']}\n"
                        f"📊 *Зрелость:* {final_score}%")
                
                # Запрос к API Telegram
                files = {'document': (f"Audit_{client_info['Наименование компании']}.xlsx", excel_file)}
                requests.post(
                    f"https://api.telegram.org/bot{TOKEN}/sendDocument",
                    data={"chat_id": CHAT_ID, "caption": text, "parse_mode": "Markdown"},
                    files=files
                )
                st.success("Отчет успешно отправлен экспертам!")
                st.balloons()
                st.download_button("📥 Скачать копию Excel", excel_file, f"Audit_{client_info['Наименование компании']}.xlsx")
            except Exception as e:
                st.error(f"Ошибка при отправке: {e}")

st.info("Khalil Audit System v2.0 | Almaty 2026")
