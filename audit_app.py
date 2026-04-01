import streamlit as st
import pandas as pd
import os
import requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 7.0a", layout="wide", page_icon="🛡️")

# Легкий якорь в начале страницы
st.anchor("top")

# --- НАСТРОЙКИ TELEGRAM (из Secrets) ---
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ЛОГОТИП И КОНТАКТЫ ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade IT Audit & Consulting")

st.markdown("### Мы поможем Вам стать лучше!")
st.divider()

st.title("📋 Опросник Технический аудит ИТ и ИБ (2026) v7.0a")

# --- ИНСТРУКЦИЯ ДЛЯ ПОЛЬЗОВАТЕЛЯ ---
with st.expander("📖 Инструкция по заполнению (нажмите, чтобы развернуть)"):
    st.markdown("""
    ### Руководство по проведению экспресс-аудита
    
    1. **Общая информация:** Укажите корректные контактные данные. Все поля со звездочкой (*) обязательны.
    2. **Заполнение блоков:** Используйте переключатели (toggles) для активации нужных подразделов.
    3. **Примечания:** Поля «Примечание» в каждом блоке **не являются обязательными**.
    4. **Логический контроль:** Сумма ОС на АРМ должна быть равна общему числу АРМ.
    5. **Результат:** Нажмите кнопку «Сформировать экспертный отчет» для получения файла Excel.
    """)

data = {}
client_info = {}
validation_errors = []
score = 0

# --- ШАПКА: ОБЩАЯ ИНФОРМАЦИЯ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город*")
    client_info['Наименование компании'] = st.text_input("Наименование компании*")
    site_input = st.text_input("Сайт компании*", key="site_field", placeholder="example.kz")
    client_info['Сайт компании'] = site_input
    
    if st.checkbox("Email отличается от сайта"):
        client_info['Email'] = st.text_input("Email контактного лица*")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица (логин)*")
            e_col1, e_col2 = st.columns([2, 3])
            with e_col1:
                email_prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="email_pre")
            with e_col2:
                st.markdown(f"<div style='padding-top: 5px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
            client_info['Email'] = f"{email_prefix}@{clean_domain}" if email_prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица*")
    client_info['Должность'] = st.text_input("Должность*")
    st.write("Контактный телефон*")
    p_col1, p_col2 = st.columns([1, 2])
    country_codes = [("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7")]
    selected_code = p_col1.selectbox("Код", country_codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed")
    client_info['Контактный телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

st.divider()

# --- БЛОК 1: ИТ ---
st.header("Блок 1: Информационные технологии")
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт)", min_value=0, step=1)
data['1.1. Всего АРМ'] = total_arm
selected_os_arm = st.multiselect("ОС на АРМ", ["Windows XP/Vista/7/8", "Windows 10", "Windows 11", "Linux", "macOS"])
sum_os_arm = 0
if selected_os_arm:
    for os_item in selected_os_arm:
        val = st.number_input(f"Кол-во на {os_item}", min_value=0, key=f"arm_{os_item}")
        data[f"ОС АРМ ({os_item})"] = val
        sum_os_arm += val
data['1.1. Примечание'] = st.text_area("Примечание к 1.1", key="note_1_1")

st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_toggle"):
    routing_opts = ["Статическая", "RIP", "OSPF", "EIGRP", "BGP", "IS-IS", "Другое"]
    sel_routing = st.multiselect("Тип маршрутизации*", routing_opts)
    data['1.2.3. Маршрутизация'] = ", ".join(sel_routing)
    
    if st.checkbox("Межсетевой экран (NGFW)"):
        v = st.text_input("Вендор NGFW*")
        data['1.2.7. NGFW'] = v
        score += 20
    data['1.2. Примечание'] = st.text_area("Примечание к 1.2", key="note_1_2")

st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
phys_srv = st.number_input("Физические серверы", min_value=0)
virt_srv = st.number_input("Виртуальные серверы", min_value=0)
data['1.3.1. Физические'] = phys_srv
data['1.3.2. Виртуальные'] = virt_srv
data['1.3. Примечание'] = st.text_area("Примечание к 1.3", key="note_1_3")

st.write("---")
st.subheader("1.5. Внутренние ИС")
if st.toggle("ИС организации", key="is_toggle"):
    is_list = {"1С": "1c", "Битрикс24": "b24", "Documentolog": "doc", "Almexoft": "alm", "HelpDeskEddy": "hde", "SAP": "sap"}
    for label, ks in is_list.items():
        if st.checkbox(label):
            ver = st.text_input(f"Версия {label}*", key=f"v_{ks}")
            data[f"ИС {label}"] = ver
    data['1.5. Примечание'] = st.text_area("Примечание к 1.5", key="note_1_5")

# --- БЛОК 2: ИБ ---
st.divider()
st.header("Блок 2: Информационная Безопасность")
if st.toggle("Средства защиты", key="ib_toggle"):
    ib_map = {"EPP": 10, "DLP": 15, "PAM": 10, "SIEM": 20, "MFA": 15}
    for label, pts in ib_map.items():
        if st.checkbox(label):
            v_ib = st.text_input(f"Вендор {label}*", key=f"ibv_{label}")
            data[label] = v_ib
            score += pts
    data['Блок 2. Примечание'] = st.text_area("Примечание к ИБ", key="note_ib")

# --- БЛОК 3: ОБЛАЧНЫЕ СЕРВИСЫ ---
st.divider()
st.header("Блок 3: Облачные сервисы")
if st.toggle("Облака", key="cloud_toggle"):
    cloud_list = ["Azure", "AWS", "Google Cloud", "Yandex Cloud", "SberCloud", "BTS Digital", "PS.kz", "Другое"]
    sel_clouds = st.multiselect("Используемые провайдеры", cloud_list)
    data['3.1. Облачные провайдеры'] = ", ".join(sel_clouds)
    data['Блок 3. Примечание'] = st.text_area("Примечание к Облакам", key="note_cloud")

# --- БЛОК 4: РАЗРАБОТКА ---
st.divider()
st.header("Блок 4: Разработка")
if st.toggle("Разработка", key="dev_toggle"):
    dev_count = st.number_input("Кол-во разработчиков*", min_value=0)
    data['4.1. Разработчики'] = dev_count
    lang_list = ["Python", "JavaScript", "Java", "C#", "Go", "PHP"]
    sel_langs = st.multiselect("Языки разработки*", lang_list)
    data['4.3. Языки'] = ", ".join(sel_langs)
    data['Блок 4. Примечание'] = st.text_area("Примечание к Разработке", key="note_dev")

# --- ГЕНЕРАЦИЯ EXCEL (Упрощенно для краткости) ---
def make_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Report"
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ 7.0a"
    row = 3
    for k, v in {**c_info, **results}.items():
        ws.cell(row=row, column=1, value=k)
        ws.cell(row=row, column=2, value=str(v))
        row += 1
    ws.cell(row=row+1, column=1, value="ИТОГО ЗРЕЛОСТЬ")
    ws.cell(row=row+1, column=2, value=f"{final_score}%")
    wb.save(output)
    return output.getvalue()

# --- ФИНАЛ ---
st.divider()
if st.button("📊 Сформировать экспертный отчет"):
    if not client_info.get('Наименование компании'):
        st.error("Заполните название компании!")
    else:
        report = make_excel(client_info, data, min(score, 100))
        st.success("Отчет готов!")
        st.download_button("📥 Скачать Excel", report, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v7.0a | Almaty 2026")
