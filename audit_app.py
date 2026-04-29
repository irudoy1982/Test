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

# --- НАСТРОЙКИ TELEGRAM (из Secrets) ---
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ЛОГОТИП ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade IT Audit & Consulting")

st.title("📋 Опросник Технический аудит ИТ и ИБ (2026) v9.2")

data = {}
client_info = {}
validation_errors = []

# --- ШАПКА: ОБЩАЯ ИНФОРМАЦИЯ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город*", help="Укажите город фактически.")
    industry_options = ["Финтех / Банки", "Ритейл / E-commerce", "Производство", "IT / Разработка", "Госсектор", "Другое"]
    selected_ind = st.selectbox("Сфера деятельности компании*", [""] + industry_options)
    industry = selected_ind if selected_ind != "Другое" else st.text_input("Укажите вашу сферу деятельности*")
    client_info['Сфера деятельности'] = industry
    client_info['Наименование компании'] = st.text_input("Наименование компании*")
    
    # ЛОГИКА САЙТА И ПОЧТЫ
    site_input = st.text_input("Сайт компании*", placeholder="example.kz")
    client_info['Сайт компании'] = site_input
    
    # Извлекаем чистый домен для почты
    domain = site_input.replace("https://", "").replace("http://", "").split("/")[0]
    email_default = f"info@{domain}" if domain else ""
    
    use_corp_email = st.checkbox("Использовать почту на домене сайта?", value=True)
    if use_corp_email:
        client_info['Email'] = st.text_input("Электронная почта (авто)", value=email_default)
    else:
        client_info['Email'] = st.text_input("Укажите другую электронную почту")

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица*")
    client_info['Должность'] = st.text_input("Должность*")
    st.write("Контактный телефон*")
    p_col1, p_col2 = st.columns([1, 2])
    selected_code = p_col1.selectbox("Код", ["+7", "+998", "+996", "+971"], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed")
    client_info['Контактный телефон'] = f"{selected_code} {phone_num}" if phone_num else ""

st.divider()

# --- БЛОК 1: ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ ---
st.header("Блок 1: Информационные технологии")

# 1.1 АРМ
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт)", min_value=0, step=1)

# 1.2 Сеть
st.subheader("1.2. Сетевая инфраструктура")
main_speed, back_speed, ap_cnt, wifi_ctrl = 0, 0, 0, False
selected_routing = []
ngfw_vendor = "Нет"
nad, nad_v = False, ""

if st.toggle("Своя сетевая инфраструктура", key="net_toggle"):
    col_net1, col_net2 = st.columns(2)
    with col_net1:
        main_speed = st.number_input("Скорость основного канала (Mbit/s)", min_value=0, step=10)
    with col_net2:
        back_speed = st.number_input("Скорость резервного канала (Mbit/s)", min_value=0, step=10)
    
    selected_routing = st.multiselect("Тип маршрутизации*", ["Статическая", "OSPF", "BGP", "EIGRP"])
    
    st.write("Беспроводная сеть")
    w_c1, w_c2 = st.columns(2)
    with w_c1:
        ap_cnt = st.number_input("Количество точек доступа (шт)", min_value=0, step=1)
    with w_c2:
        wifi_ctrl = st.checkbox("Наличие Wi-Fi контроллера")

    st.markdown("**Сетевая безопасность**")
    if st.checkbox("Межсетевой экран (NGFW)"):
        ngfw_vendor = st.text_input("Производитель NGFW")
    
    nad = st.checkbox("NAD (Network Analysis and Detection)", help="Перенесено в сетевую безопасность")
    nad_v = st.text_input("Производитель NAD", key="nad_v_ui") if nad else ""

# 1.3 Серверы, СХД, ИС
col_s1, col_s2 = st.columns(2)
with col_s1:
    st.subheader("1.3. Серверы")
    phys_count = st.number_input("Физические серверы", min_value=0)
    virt_count = st.number_input("Виртуальные серверы", min_value=0)
    v_n_b = st.text_input("Вендор СРК (Резервное копирование)") if st.checkbox("Есть СРК") else "Нет"

with col_s2:
    st.subheader("1.4. Хранение (СХД)")
    storage_data = "Нет"
    if st.toggle("Есть СХД"):
        st_vendor = st.text_input("Вендор СХД")
        st_raid = st.multiselect("Уровни RAID", ["RAID 0", "RAID 1", "RAID 5", "RAID 6", "RAID 10"])
        storage_data = f"{st_vendor} (RAID: {', '.join(st_raid)})"

st.subheader("1.5. Информационные системы")
helpdesk_exists = False
if st.toggle("Бизнес-приложения"):
    if st.checkbox("HelpDesk / ServiceDesk"):
        hd_name = st.text_input("Название системы (например, Jira, GLPI)")
        helpdesk_exists = True

st.divider()

# --- БЛОК 2-4: ИБ, WEB, DEV ---
col_ib, col_web_dev = st.columns(2)

with col_ib:
    st.header("Блок 2: ИБ")
    ib_list = ["EPP", "EDR", "DLP", "SIEM", "PAM", "MFA", "WAF"]
    ib_values = {}
    for item in ib_list:
        if st.checkbox(item):
            ib_values[item] = st.text_input(f"Вендор {item}", key=f"v_{item}")
        else:
            ib_values[item] = "Нет"

with col_web_dev:
    st.header("Блок 3-4: Web & Dev")
    web_status = st.text_input("Web-стек (Nginx/IIS/Cloud)") if st.toggle("Web-ресурсы") else "Нет"
    dev_status = st.text_input("Стек разработки") if st.toggle("Собственная разработка") else "Нет"

# --- ЛОГИКА ОТЧЕТА ---
def make_expert_excel(c_info, results):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Report 2026"

    # Стилизация
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    
    # 1. Шапка клиента
    ws['A1'] = "ДАННЫЕ КЛИЕНТА"
    ws['A1'].font = Font(bold=True, size=12)
    curr_row = 2
    for k, v in c_info.items():
        ws.cell(row=curr_row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=curr_row, column=2, value=str(v))
        curr_row += 1
    
    curr_row += 1
    ws.cell(row=curr_row, column=1, value="КРИТИЧЕСКИЙ АНАЛИЗ (CTO/CISO)").font = Font(bold=True, color="FF0000")
    curr_row += 1

    # ЛОГИКА КАНАЛОВ
    ms = results.get("Интернет скорость", 0)
    bs = results.get("Резерв скорость", 0)
    if bs == 0:
        ws.cell(row=curr_row, column=1, value="• Отсутствие резервного канала — единая точка отказа."); curr_row += 1
    elif bs < (ms / 2):
        ws.cell(row=curr_row, column=1, value="• ВНИМАНИЕ: Резервный канал не покрывает производительность основного (менее 50%)."); curr_row += 1

    # ЛОГИКА WI-FI
    ap = results.get("WiFi Точки", 0)
    users = results.get("Пользователи", 0)
    if ap > 0:
        ratio = users / ap
        if ratio > 25:
            ws.cell(row=curr_row, column=1, value=f"• Перегрузка Wi-Fi: {ratio:.1f} чел. на точку. Нужно расширение."); curr_row += 1
        if ap > 5 and results.get("WiFi Контроллер") == "Нет":
            ws.cell(row=curr_row, column=1, value="• Проблема: Более 5 ТД без контроллера — риски безопасности и управления."); curr_row += 1

    # ЛОГИКА ПРОДУКТОВ (SIEM / HelpDesk)
    total_hosts = results.get("Серверы", 0) + results.get("Виртуализация", 0) + users
    if total_hosts > 50 and results.get("SIEM") == "Нет":
        ws.cell(row=curr_row, column=1, value="• Рекомендация CISO: Внедрить SIEM-систему для мониторинга событий безопасности."); curr_row += 1
    
    if users > 30 and results.get("HelpDesk") == "Нет":
        ws.cell(row=curr_row, column=1, value="• Рекомендация CTO: Необходима HelpDesk система для автоматизации заявок."); curr_row += 1

    curr_row += 2
    # Детальная таблица всех блоков
    headers = ["Параметр", "Значение", "Статус"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=curr_row, column=i, value=h).fill = header_fill
        ws.cell(row=curr_row, column=i).font = white_font
    curr_row += 1

    for k, v in results.items():
        ws.cell(row=curr_row, column=1, value=k)
        ws.cell(row=curr_row, column=2, value=str(v))
        ws.cell(row=curr_row, column=3, value="OK" if "Нет" not in str(v) and v != 0 else "Внимание")
        curr_row += 1

    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 40
    wb.save(output)
    return output.getvalue()

# --- КНОПКА ОТПРАВКИ ---
if st.button("🚀 Сформировать и отправить отчет"):
    all_results = {
        "Пользователи": total_arm,
        "Интернет скорость": main_speed,
        "Резерв скорость": back_speed,
        "WiFi Точки": ap_cnt,
        "WiFi Контроллер": "Да" if wifi_ctrl else "Нет",
        "NGFW": ngfw_vendor,
        "NAD": nad_v if nad else "Нет",
        "Серверы": phys_count,
        "Виртуализация": virt_count,
        "Резервное копирование": v_n_b,
        "СХД": storage_data,
        "HelpDesk": "Да" if helpdesk_exists else "Нет",
        "Web-ресурсы": web_status,
        "Разработка": dev_status
    }
    all_results.update(ib_values)

    report = make_expert_excel(client_info, all_results)
    
    # Отправка в ТГ (тихая)
    try:
        requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                      data={"chat_id": CHAT_ID, "caption": f"Новый аудит: {client_info['Наименование компании']}"}, 
                      files={'document': (f"Report_{client_info['Наименование компании']}.xlsx", report)})
    except: pass

    st.success("Отчет готов! Все блоки (СХД, Web, Dev) и рекомендации CTO/CISO включены.")
    st.download_button("📥 Скачать Excel", report, f"Audit_{client_info['Наименование компании']}.xlsx")
