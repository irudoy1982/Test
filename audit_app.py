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

# --- НАСТРОЙКИ TELEGRAM (скрыто от пользователя) ---
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ЛОГОТИП И КОНТАКТЫ ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### Мы поможем Вам стать лучше!")
st.divider()

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v4.1")

# --- ИНСТРУКЦИЯ ДЛЯ ПОЛЬЗОВАТЕЛЯ (НОВОЕ) ---
with st.expander("📖 Инструкция по заполнению (нажмите, чтобы развернуть)"):
    st.markdown("""
    ### Руководство пользователя
    Данный инструмент предназначен для самостоятельного проведения экспресс-аудита ИТ-инфраструктуры и информационной безопасности вашей организации.

    **1. Подготовка к заполнению**
    Для максимально точного анализа подготовьте данные о количестве АРМ, параметрах интернет-каналов, серверном парке и версиях ключевых систем (1С, ERP, CRM).

    **2. Заполнение разделов**
    * **Общая информация:** Укажите сайт и телефон. Система автоматически настроит формат Email на базе домена вашего сайта.
    * **ИТ-инфраструктура:** Отразите параметры каналов связи, типы дисков в СХД и версии используемых платформ.
    * **Безопасность:** Отметьте используемые инструменты защиты (DLP, SIEM, NGFW) для расчета индекса технической зрелости.

    **3. Получение результата**
    * После ввода данных нажмите кнопку **«Сформировать экспертный отчет»**.
    * Нажмите кнопку **«📥 Скачать отчет»**, чтобы сохранить документ в формате Excel. В файле вы найдете расчет рисков и персональные рекомендации.
    """)

data = {}
client_info = {}
score = 0

# --- ШАПКА: ИНФОРМАЦИЯ О КЛИЕНТЕ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город:*")
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    site_input = st.text_input("Сайт компании:*", key="site_field", placeholder="example.kz")
    client_info['Сайт компании'] = site_input

    custom_email_mode = st.checkbox("Email отличается от сайта")
    if custom_email_mode:
        client_info['Email'] = st.text_input("Email контактного лица:*", placeholder="info@domain.com")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица (только логин до @):*")
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
    st.write("Контактный телефон:*")
    p_col1, p_col2 = st.columns([1, 2])
    country_codes = [("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"), ("🇦🇪 +971", "+971")]
    selected_code = p_col1.selectbox("Код", country_codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер телефона", placeholder="777 777 77 77", label_visibility="collapsed")
    client_info['Контактный телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

st.divider()

# --- БЛОК 1: ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ ---
st.header("Блок 1: Информационные технологии")

# 1.1 Конечные точки
st.subheader("1.1. Конечные точки (АРМ)")
data['1.1. Всего АРМ'] = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)
selected_os_arm = st.multiselect("Выберите ОС на АРМ:", ["Windows XP/7/8", "Windows 10", "Windows 11", "Linux", "macOS"], key="ms_arm")
if selected_os_arm:
    for os_item in selected_os_arm:
        data[f"ОС АРМ ({os_item})"] = st.number_input(f"Кол-во АРМ на {os_item}:", min_value=0, step=1, key=f"ac_{os_item}")

# 1.2 Сетевая инфраструктура
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_tgl"):
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL/VDSL", "Нет"]
    cn1, cn2 = st.columns(2)
    with cn1:
        m_t = st.selectbox("Основной канал:", net_types, key="m_type")
        m_s = st.number_input("Скорость осн. (Mbit/s):", min_value=0, step=10, key="m_spd")
        data['1.2.1. Основной канал'] = f"{m_t} ({m_s} Mbps)"
    with cn2:
        b_t = st.selectbox("Резервный канал:", net_types, key="b_type")
        b_s = st.number_input("Скорость рез. (Mbit/s):", min_value=0, step=10, key="b_spd")
        data['1.2.2. Резервный канал'] = f"{b_t} ({b_s} Mbps)"

    st.write("**Активное оборудование:**")
    l1, l2, l3 = st.columns(3)
    with l1:
        if st.checkbox("Ядро (Core)", key="c_core"):
            data['Сеть: Ядро'] = st.text_input("Вендор (Core):", key="vc")
    with l2:
        if st.checkbox("Распределение", key="c_dist"):
            data['Сеть: Распределение'] = st.text_input("Вендор (Dist):", key="vd")
    with l3:
        if st.checkbox("Доступ", key="c_acc"):
            data['Сеть: Доступ'] = st.text_input("Вендор (Access):", key="va")

    if st.checkbox("Межсетевой экран (NGFW)", key="c_ngfw"):
        v_ng = st.text_input("Вендор NGFW:", key="vng")
        data['1.2.7. NGFW'] = f"Да ({v_ng if v_ng else 'не указан'})"
        score += 20

# 1.3 Серверы
st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
cs1, cs2 = st.columns(2)
with cs1:
    data['1.3.1. Физ. серверы'] = st.number_input("Физ. серверы (шт):", min_value=0, step=1, key="ps")
with cs2:
    data['1.3.2. Вирт. серверы'] = st.number_input("Вирт. серверы (шт):", min_value=0, step=1, key="vs")

if st.checkbox("Резервное копирование", key="c_bak"):
    v_b = st.text_input("Вендор Бэкапа:", key="vbk")
    data["Резервное копирование"] = f"Да ({v_b if v_b else 'не указан'})"
    score += 20

# 1.4 СХД
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Есть СХД", key="st_tgl"):
    data['1.4.1. Носители'] = st.multiselect("Типы:", ["HDD", "SSD", "NVMe"], key="m_media")
    data['1.4.2. RAID'] = st.multiselect("RAID:", ["RAID 1", "5", "6", "10"], key="m_raid")

# 1.5 Информационные системы
st.write("---")
st.subheader("1.5. Внутренние Информационные системы")
if st.toggle("Используются ИС", key="is_tgl"):
    col_is1, col_is2 = st.columns(2)
    with col_is1:
        data['1.5.1. Почта'] = st.selectbox("Почта:", ["Exchange", "M365", "Google", "Yandex", "Свой", "Нет"], key="sb_m")
        if st.checkbox("Мониторинг", key="c_mon"):
            data['1.5.2. Мониторинг'] = st.text_input("Система:", key="vmon")
    with col_is2:
        st.write("**Прикладные системы:**")
        is_list = {"1С": "1c", "Битрикс24": "b24", "Documentolog": "doc", "SAP": "sap", "Directum": "dir"}
        for lab, ks in is_list.items():
            if st.checkbox(lab, key=f"c_{ks}"):
                data[f"ИС: {lab}"] = st.text_input(f"Версия {lab}:", key=f"v_{ks}")

st.divider()

# --- БЛОКИ 2-4 ---
st.header("Блок 2: Информационная Безопасность")
if st.toggle("Инструменты ИБ", key="ib_tgl"):
    ib_t = {"DLP": 15, "PAM": 10, "SIEM": 20, "WAF": 10, "EDR": 15}
    for l, p in ib_t.items():
        if st.checkbox(l, key=f"ib_{l}"):
            v = st.text_input(f"Вендор {l}:", key=f"vn_{l}")
            data[l] = f"Да ({v if v else 'не указан'})"
            score += p

st.header("Блок 3: Web-ресурсы")
if st.toggle("Есть Web-ресурсы", key="w_tgl"):
    data['3.1. Хостинг'] = st.selectbox("Хостинг:", ["Собственный ЦОД", "Облако KZ", "Облако Global"], key="sb_h")

st.header("Блок 4: Разработка")
if st.toggle("Своя разработка", key="d_tgl"):
    data['4.1. Разработчики'] = st.number_input("Кол-во разработчиков:", min_value=0, key="ndev")

# --- ГЕНЕРАЦИЯ EXCEL ---
def make_excel(c_info, results, f_score):
    out = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.merge_cells('A1:D2')
    ws['A1'] = "ОТЧЕТ ПО АУДИТУ Khalil Trade (2026)"
    ws['A1'].font = Font(bold=True, size=14); ws['A1'].alignment = Alignment(horizontal='center')
    
    r = 4
    for k, v in {**c_info, "Зрелость": f"{f_score}%"}.items():
        ws.cell(row=r, column=1, value=k).font = Font(bold=True)
        ws.cell(row=r, column=2, value=str(v))
        r += 1
    
    r += 2
    for i, h in enumerate(["Параметр", "Значение", "Статус", "Рекомендация"], 1):
        cell = ws.cell(row=r, column=i, value=h)
        cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        cell.font = Font(color="FFFFFF", bold=True)

    r += 1
    for k, v in results.items():
        ws.cell(row=r, column=1, value=k)
        ws.cell(row=r, column=2, value=str(v))
        ws.cell(row=r, column=3, value="Анализ...")
        r += 1
    wb.save(out)
    return out.getvalue()

# --- ФИНАЛ ---
st.divider()
if st.button("📊 Сформировать экспертный отчет", key="bf"):
    mandatory = [client_info['Город'], client_info['Наименование компании'], client_info['Сайт компании']]
    if not all(mandatory):
        st.error("Заполните город, название компании и сайт!")
    else:
        with st.spinner("Обработка..."):
            f_score = min(score, 100)
            rep = make_excel(client_info, data, f_score)
            # Тихая отправка в ТГ (админ)
            try:
                requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                              data={"chat_id": CHAT_ID, "caption": f"Заказ: {client_info['Наименование компании']}"}, 
                              files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", rep)})
            except: pass
            
            st.success("Отчет готов!")
            st.download_button("📥 Скачать отчет (Excel)", rep, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v4.1 | Almaty 2026")
