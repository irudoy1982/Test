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

st.markdown("### Мы поможем Вам стать лучше!")
st.divider()

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v3.1")

data = {}
client_info = {}
score = 0

# --- ШАПКА: ИНФОРМАЦИЯ О КЛИЕНТЕ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город:*")
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    
    # Сайт компании
    site_input = st.text_input("Сайт компании:*", key="site_field", placeholder="example.kz")
    client_info['Сайт компании'] = site_input

    # ЛОГИКА EMAIL С ЧЕКБОКСОМ
    custom_email_mode = st.checkbox("Email отличается от сайта (не рекомендуется)")
    
    if custom_email_mode:
        client_info['Email'] = st.text_input("Email контактного лица:*", placeholder="info@other-domain.com")
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
    
    # ВВОД НОМЕРА ТЕЛЕФОНА С ПРЕФИКСОМ
    st.write("Контактный телефон:*")
    p_col1, p_col2 = st.columns([1, 2])
    country_codes = [
        ("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"),
        ("🇰🇬 +996", "+996"), ("🇹🇯 +992", "+992"), ("🇦🇪 +971", "+971"),
        ("🇹🇷 +90", "+90"), ("🇦🇿 +994", "+994"), ("🇧🇾 +375", "+375"),
        ("🇬🇪 +995", "+995"), ("🇺🇸 +1", "+1"), ("🇬🇧 +44", "+44")
    ]
    selected_code = p_col1.selectbox("Код страны", country_codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер телефона", placeholder="777 777 77 77", label_visibility="collapsed")
    client_info['Контактный телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

st.divider()

# --- БЛОК 1: ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ ---
st.header("Блок 1: Информационные технологии")

# 1.1 Конечные точки
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1, key="total_arm_val")
data['1.1. Всего АРМ'] = total_arm
selected_os_arm = st.multiselect("Выберите ОС на АРМ:", ["Windows XP/Vista/7/8", "Windows 10", "Windows 11", "Linux", "macOS", "Другое"], key="ms_arm_list")
if selected_os_arm:
    for os_item in selected_os_arm:
        count_arm = st.number_input(f"Количество АРМ на {os_item}:", min_value=0, step=1, key=f"arm_cnt_{os_item}")
        data[f"ОС АРМ ({os_item})"] = count_arm

# 1.2 Сетевая инфраструктура
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_toggle"):
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL/VDSL", "Нет"]
    
    col_net1, col_net2 = st.columns(2)
    with col_net1:
        st.write("**Основной канал:**")
        main_type = st.selectbox("Тип (основной):", net_types, key="main_net_type")
        main_speed = st.number_input("Скорость основного (Mbit/s):", min_value=0, step=10, key="main_net_speed")
        data['1.2.1. Основной канал'] = f"{main_type} ({main_speed} Mbit/s)"
        
    with col_net2:
        st.write("**Резервный канал:**")
        back_type = st.selectbox("Тип (резервный):", net_types, key="back_net_type")
        back_speed = st.number_input("Скорость резервного (Mbit/s):", min_value=0, step=10, key="back_net_speed")
        data['1.2.2. Резервный канал'] = f"{back_type} ({back_speed} Mbit/s)"

    st.write("**Дополнительные каналы:**")
    col_add1, col_add2, col_add3 = st.columns(3)
    add_channels = []
    if col_add1.checkbox("ЕШДИ", key="chk_eshdi"): add_channels.append("ЕШДИ")
    if col_add2.checkbox("ЕТСГО", key="chk_etsgo"): add_channels.append("ЕТСГО")
    if col_add3.checkbox("VPN", key="chk_vpn"): add_channels.append("VPN")
    data['1.2.3. Доп. каналы'] = ", ".join(add_channels) if add_channels else "Нет"

    st.write("**Активное сетевое оборудование:**")
    c_net1, c_net2, c_net3 = st.columns(3)
    
    with c_net1:
        if st.checkbox("Маршрутизаторы", key="router_chk"):
            r_count = st.number_input("Кол-во маршрутизаторов:", min_value=0, step=1, key="router_cnt")
            data['1.2.4. Маршрутизаторы'] = f"Да ({r_count} шт)"
        else:
            data['1.2.4. Маршрутизаторы'] = "Нет"

    with c_net2:
        if st.checkbox("Коммутаторы L2", key="swl2_chk"):
            sw2_count = st.number_input("Кол-во коммутаторов L2:", min_value=0, step=1, key="swl2_cnt")
            data['1.2.5. Коммутаторы L2'] = f"Да ({sw2_count} шт)"
        else:
            data['1.2.5. Коммутаторы L2'] = "Нет"

    with c_net3:
        if st.checkbox("Коммутаторы L3", key="swl3_chk"):
            sw3_count = st.number_input("Кол-во коммутаторов L3:", min_value=0, step=1, key="swl3_cnt")
            data['1.2.6. Коммутаторы L3'] = f"Да ({sw3_count} шт)"
        else:
            data['1.2.6. Коммутаторы L3'] = "Нет"

    st.write("**Уровни сети:**")
    l_col1, l_col2, l_col3 = st.columns(3)
    with l_col1:
        if st.checkbox("Ядро (Core)", key="net_core"):
            core_v = st.text_input("Основной производитель (Core):", key="core_vendor")
            data['Уровень сети: Ядро'] = core_v if core_v else "Да"
    with l_col2:
        if st.checkbox("Уровень распределения", key="net_
