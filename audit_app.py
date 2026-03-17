import streamlit as st
import pandas as pd
import os
import requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

# --- ПОЛУЧЕНИЕ НАСТРОЕК TELEGRAM ИЗ SECRETS ---
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
score = 0

# --- БЛОК 1: ОБЩАЯ ИНФОРМАЦИЯ ---
st.header("Блок 1: Общая информация")

# 1.1 Конечные точки (АРМ)
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1, key="total_arm_val")
data['1.1. Всего АРМ'] = total_arm

selected_os_arm = st.multiselect("Выберите ОС на АРМ:", ["Windows", "Linux", "macOS", "Другое"], key="ms_arm_list")
if selected_os_arm:
    for os_item in selected_os_arm:
        count_arm = st.number_input(f"Количество АРМ на {os_item}:", min_value=0, step=1, key=f"arm_cnt_{os_item}")
        data[f"ОС АРМ ({os_item})"] = count_arm

st.write("---")

# 1.2 Серверная инфраструктура
st.subheader("1.2. Серверы")
col_s1, col_s2 = st.columns(2)
with col_s1:
    phys_servers = st.number_input("Количество физических серверов:", min_value=0, step=1, key="phys_srv")
    data['1.2. Физические серверы'] = phys_servers
with col_s2:
    virt_servers = st.number_input("Количество виртуальных серверов:", min_value=0, step=1, key="virt_srv")
    data['1.2. Виртуальные серверы'] = virt_servers

selected_os_srv = st.multiselect("Выберите ОС серверов:", ["Windows Server", "Linux", "Unix", "Другое"], key="ms_srv_list")
if selected_os_srv:
    for os_s in selected_os_srv:
        count_srv = st.number_input(f"Количество серверов на {os_s}:", min_value=0, step=1, key=f"srv_cnt_{os_s}")
        data[f"ОС Сервера ({os_s})"] = count_srv

st.write("---")

# 1.3 Виртуализация и 1.4 Почта
col_v1, col_v2 = st.columns(2)
with col_v1:
    st.subheader("1.3. Виртуализация")
    data['1.3. Виртуализация'] = st.multiselect("Системы виртуализации:", ["VMware", "Hyper-V", "Proxmox", "KVM", "Другое", "Нет"], key="virt_sys")
with col_v2:
    st.subheader("1.4. Почтовая система")
    data['1.4. Почта'] = st.selectbox("Тип почты:", ["Exchange (On-Prem)", "Microsoft 365", "Google Workspace", "Yandex/Mail.ru Cloud", "Собственный сервер", "Нет"], key="mail_sys")

st.write("---")

# 1.5 Информационные системы
st.subheader("1.5. Внутренние Информационные системы")
has_is = st.checkbox("Есть ли внутренние Информационные системы (1C, ERP, CRM)?", key="is_15_chk")
if has_is:
    data['1.5. Внутренние ИС'] = st.text_input("Перечислите их через запятую:", key="is_input_field")
else:
    data['1.5. Внутренние ИС'] = "Нет"

# 1.6 Мониторинг
st.subheader("1.6. Система мониторинга")
has_mon = st.checkbox("Есть ли система мониторинга?", key="mon_16_chk")
if has_mon:
    data['1.6. Мониторинг'] = st.selectbox("Выберите систему:", ["Zabbix", "Nagios", "PRTG", "Prometheus", "Другое"], key="mon_select_field")
else:
    data['1.6. Мониторинг'] = "Нет"


# --- БЛОК 2: СЕТЬ ---
st.header("Блок 2: Сетевая инфраструктура и Интернет")
if st.toggle("Своя сетевая инфраструктура", key="net_block_toggle"):
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink"]
    
    c_n1, c_n2 = st.columns(2)
    with c_n1:
        data['2.1. Основной канал'] = st.selectbox("Тип канала (Основной):", net_types, key="main_net_type")
        data['2.1. Скорость осн. (mbit/s)'] = st.number_input("Скорость основного канала (mbit/s):", min_value=0, key="main_net_speed")
    with c_n2:
        data['2.1. Резервный канал'] = st.selectbox("Тип канала (Резервный):", ["Нет"] + net_types, key="res_net_type")
        data['2.1. Скорость рез. (mbit/s)'] = st.number_input("Скорость резервного канала (mbit/s):", min_value=0, key="res_net_speed")

    data['2.4. NGFW'] = st.text_input("Вендор Межсетевого экрана (NGFW):", key="ngfw_vendor_input")
    if data['2.4. NGFW']: score += 20

    has_wifi = st.checkbox("Используется Wi-Fi?", key="wifi_usage_chk")
    if has_wifi:
        if st.checkbox("Есть ли Wi-Fi контроллер?", key="wifi_ctrl_exists"):
            data['2.5. Контроллер'] = st.text_input("Модель контроллера:", key="wifi_ctrl_model")
        else:
            data['2.5. Контроллер'] = "Без контроллера"
        data['2.5. Число точек'] = st.number_input("Кол-во точек доступа:", min_value=0, key="wifi_ap_count")


# --- БЛОК 3: ИБ ---
st.header("Блок 3: Информационная Безопасность")
if st.toggle("Есть отдел Информационной безопасности", key="ib_block_main_toggle"):
    st.info("Отметьте внедренные системы и укажите их вендора:")
    ib_list = {
        "DLP (Защита от утечек)": 15,
        "PAM (Контроль доступа)": 10,
        "SIEM/SOC (Мониторинг ИБ)": 20,
        "WAF (Защита Web)": 10,
        "EDR/Antimalware": 15,
        "Резервное копирование": 20
    }
    for label, pts in ib_list.items():
        c_ib1, c_ib2 = st.columns([1, 2])
        with c_ib1:
            is_on = st.checkbox(label, key=f"ib_v2_{label}")
        if is_on:
            with c_ib2:
                v_name = st.text_input(f"Вендор {label}:", key=f"v_name_{label}")
                data[label] = f"Да ({v_name if v_name else 'не указан'})"
                score += pts
        else:
            data[label] = "Нет"
    
    st.write("---")
    if st.checkbox("3.7. Другое (дополнительные системы защиты)", key="ib_other_toggle"):
        data['3.7. Прочие системы ИБ'] = st.text_area("Перечислите все, что мы не учли:", key="ib_other_input")
    else:
        data['3
