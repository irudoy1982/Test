import streamlit as st
import pandas as pd
import os
import requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

# Данные для Telegram из Secrets (настраиваются в Streamlit Cloud)
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
    client_info['Компания'] = st.text_input("Наименование компании:*")
    site_in = st.text_input("Сайт:*", placeholder="example.kz")
    client_info['Сайт'] = site_in

    if st.checkbox("Email отличается от домена сайта"):
        client_info['Email'] = st.text_input("Email:*", placeholder="info@domain.com")
    else:
        domain = site_in.replace("https://","").replace("http://","").replace("www.","").split('/')[0]
        if domain and "." in domain:
            st.write("Email (логин):*")
            e1, e2 = st.columns([2, 3])
            prefix = e1.text_input("Login", placeholder="info", label_visibility="collapsed", key="em_pre")
            e2.markdown(f"**@{domain}**")
            client_info['Email'] = f"{prefix}@{domain}" if prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    st.write("Телефон:*")
    p1, p2 = st.columns([1, 2])
    codes = [
        ("🇰🇿 +7","+7"), ("🇷🇺 +7","+7"), ("🇺🇿 +998","+998"), 
        ("🇰🇬 +996","+996"), ("🇹🇯 +992","+992"), ("🇦🇪 +971","+971"),
        ("🇹🇷 +90","+90"), ("🇦🇿 +994","+994"), ("🇧🇾 +375","+375"),
        ("🇬🇪 +995","+995"), ("🇺🇸 +1","+1"), ("🇬🇧 +44","+44")
    ]
    code = p1.selectbox("Код", codes, format_func=lambda x: x[0], label_visibility="collapsed")
    num = p2.text_input("Номер", placeholder="701 123 45 67", label_visibility="collapsed", key="ph_num")
    client_info['Телефон'] = f"{code[1]} {num}"

st.divider()

# --- БЛОК 1: ИТ ИНФРАСТРУКТУРА ---
st.header("Блок 1: ИТ Инфраструктура")
data['АРМ (всего)'] = st.number_input("Кол-во АРМ (шт):", min_value=0, step=1)

if st.toggle("Своя сетевая инфраструктура"):
    data['Канал'] = st.selectbox("Тип связи:", ["Оптика", "Радио", "Спутник", "4G/5G", "Starlink"])
    data['NGFW'] = st.text_input("Вендор Межсетевого экрана (NGFW):")
    if data['NGFW']: score += 20

data['Серверы (физ)'] = st.number_input("Физические серверы:", min_value=0)
data['Серверы (вирт)'] = st.number_input("Виртуальные серверы:", min_value=0)
data['Почтовая система'] = st.selectbox("Почта:", ["Exchange", "M365", "Google", "Yandex", "Свой", "Нет"])

# --- БЛОК 2: ИБ (ТУТ БЫЛА ОШИБКА) ---
st.header("Блок 2: Информационная безопасность")
data['Антивирус'] = st.text_input("Используемое антивирусное ПО (EDR/AV):")
if data['Антивирус']: score += 15

data['Бэкапы'] = st.radio("Наличие резервного копирования данных:", ["Да", "Нет", "Частично"])
if data['Бэкапы'] == "Да": score += 25

data['VPN'] = st.checkbox("Используется ли VPN для удаленного доступа?")
if data['VPN']: score += 10

st.divider()

# --- ФИНАЛЬНАЯ ЛОГИКА ---
if st.button("🚀 Отправить данные на аудит"):
    # Валидация обязательных полей
    if not client_info['Компания'] or not client_info['ФИО'] or not num:
        st.error("Пожалуйста, заполните обязательные поля: Компания, ФИО и Номер телефона.")
    else:
        # 1. Создание Excel в памяти
        wb = Workbook()
        ws = wb.active
        ws.title = "Результаты аудита"
        
        # Стили
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        
        # Данные клиента
        ws.append(["Параметр", "Значение"])
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            
        for key, value in client_info.items():
            ws.append([key, value])
            
        ws.append([]) # Разделитель
        
        # Технические данные
        ws.append(["Технический параметр", "Ответ"])
        last_row = ws.max_row
        for cell in ws[last_row]:
            cell.font = header_font
            cell.fill = header_fill
            
        for key, value in data.items():
            ws.append([key, str(value)])
            
        ws.append([])
        ws.append(["ИТОГОВЫЙ БАЛЛ БЕЗОПАСНОСТИ", score])
        ws.cell(row=ws.max_row, column=2).font = Font(bold=True)

        # Сохранение в буфер
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        excel_bytes = output.getvalue()

        # 2. Отправка в Telegram
        if TOKEN and CHAT_ID:
            try:
                # Текст сообщения
                text_msg = (f"🔔 **Новая заявка на аудит!**\n\n"
                            f"🏢 Компания: {client_info['Компания']}\n"
                            f"👤 Контакт: {client_info['ФИО']}\n"
                            f"📞 Тел: {client_info['Телефон']}\n"
                            f"🛡️ Балл ИБ: {score}")
                
                requests.post(f"https://api.telegram.org/bot{TOKEN}/sendMessage", 
                             data={"chat_id": CHAT_ID, "text": text_msg, "parse_mode": "Markdown"})
                
                # Файл
                files = {'document': (f"Audit_{client_info['Компания']}.xlsx", excel_bytes)}
                requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                             data={"chat_id": CHAT_ID}, files=files)
                
                st.success("✅ Отчет успешно сформирован и отправлен в отдел аудита!")
                st.balloons()
            except Exception as e:
                st.error(f"Ошибка при отправке в Telegram: {e}")
        else:
            st.warning("⚠️ Настройки Telegram не найдены. Вы можете скачать файл вручную ниже.")

        # 3. Кнопка скачивания для пользователя
        st.download_button(
            label="📥 Скачать отчет (Excel)",
            data=excel_bytes,
            file_name=f"Audit_{client_info['Компания']}_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
