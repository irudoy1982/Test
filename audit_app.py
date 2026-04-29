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

# Якорь для принудительного перехода в начало страницы
st.markdown("<div id='top'></div>", unsafe_allow_html=True)

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

st.title("📋 Опросник Технический аудит ИТ и ИБ (2026) v9.1")

data = {}
client_info = {}
validation_errors = []
score = 0

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
    site_input = st.text_input("Сайт компании*", placeholder="example.kz")
    client_info['Сайт компании'] = site_input

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица*")
    client_info['Должность'] = st.text_input("Должность*")
    st.write("Контактный телефон*")
    p_col1, p_col2 = st.columns([1, 2])
    selected_code = p_col1.selectbox("Код", ["+7", "+998", "+996", "+971"], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed")
    client_info['Контактный телефон'] = f"{selected_code} {phone_num}" if phone_num else ""

if not all([client_info.get('Город'), client_info.get('Наименование компании'), site_input, phone_num]):
    validation_errors.append("Заполните обязательные поля в блоке 'Общая информация'")

st.divider()

# --- БЛОК 1: ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ ---
st.header("Блок 1: Информационные технологии")

# 1.1 АРМ
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт)", min_value=0, step=1)
data['1.1. Всего АРМ'] = total_arm

# 1.2 Сеть
st.subheader("1.2. Сетевая инфраструктура")
main_speed, back_speed, ap_cnt, wifi_ctrl = 0, 0, 0, False
selected_routing = []
ngfw_vendor = "Нет"
nad, nad_v = False, ""

if st.toggle("Своя сетевая инфраструктура", key="net_toggle"):
    col_net1, col_net2 = st.columns(2)
    with col_net1:
        st.write("Основной канал")
        main_speed = st.number_input("Скорость основного (Mbit/s)", min_value=0, step=10, key="ms_in")
    with col_net2:
        st.write("Резервный канал")
        back_speed = st.number_input("Скорость резервного (Mbit/s)", min_value=0, step=10, key="bs_in")
    
    selected_routing = st.multiselect("Тип маршрутизации*", ["Статическая", "OSPF", "BGP", "EIGRP"])
    
    st.write("Беспроводная сеть")
    w_c1, w_c2 = st.columns(2)
    with w_c1:
        ap_cnt = st.number_input("Количество точек доступа (шт)", min_value=0, step=1)
    with w_c2:
        wifi_ctrl = st.checkbox("Наличие Wi-Fi контроллера")

    st.write("Сетевая безопасность (L3/L4+)")
    if st.checkbox("Межсетевой экран (NGFW)"):
        ngfw_vendor = st.text_input("Производитель NGFW")
    
    nad = st.checkbox("NAD (Network Attack Discovery)", help="Система глубокого анализа сетевого трафика")
    nad_v = st.text_input("Производитель NAD", key="nad_v_ui") if nad else ""

# 1.3 Серверы
st.subheader("1.3. Серверы и Виртуализация")
s_c1, s_c2 = st.columns(2)
phys_count = s_c1.number_input("Физические серверы", min_value=0)
virt_count = s_c2.number_input("Виртуальные серверы", min_value=0)
v_n_b = "Нет"
if st.checkbox("Резервное копирование"):
    v_n_b = st.text_input("Вендор СРК")

# 1.4 СХД
st.subheader("1.4. СХД")
storage_data = "Нет"
if st.toggle("Есть собственная СХД", key="st_toggle"):
    st_vendor = st.text_input("Производитель СХД")
    st_raid = st.multiselect("RAID", ["RAID 0", "RAID 1", "RAID 5", "RAID 6", "RAID 10"])
    storage_data = f"{st_vendor} (RAID: {', '.join(st_raid)})"

# 1.5 ИС
st.subheader("1.5. Информационные системы")
helpdesk_exists = False
if st.toggle("Бизнес-приложения"):
    if st.checkbox("HelpDesk / ServiceDesk"):
        hd_name = st.text_input("Название системы HelpDesk")
        data["ИС HelpDesk"] = hd_name
        helpdesk_exists = True

st.divider()

# --- БЛОК 2: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ ---
st.header("Блок 2: Информационная безопасность")
enable_security = st.toggle("Включить детальный блок ИБ", value=False)
ib_params = {
    "EPP": (st.checkbox("EPP"), "epp_v"), "EDR": (st.checkbox("EDR"), "edr_v"),
    "DLP": (st.checkbox("DLP"), "dlp_v"), "SIEM": (st.checkbox("SIEM"), "siem_v"),
    "PAM": (st.checkbox("PAM"), "pam_v"), "MFA": (st.checkbox("MFA"), "mfa_v"),
    "WAF": (st.checkbox("WAF"), "waf_v"), "VULN": (st.checkbox("Сканер"), "vuln_v")
}
ib_values = {}
if enable_security:
    for key, (checked, v_key) in ib_params.items():
        ib_values[key] = st.text_input(f"Вендор {key}", key=v_key) if checked else "Нет"

# --- БЛОК 3: WEB & DEV ---
col_web, col_dev = st.columns(2)
with col_web:
    st.header("Блок 3: Web")
    web_status = "Нет"
    if st.toggle("Web-ресурсы"):
        web_status = st.text_input("Стек (Nginx, Apache и т.д.)")
with col_dev:
    st.header("Блок 4: Разработка")
    dev_status = "Нет"
    if st.toggle("Собственная разработка"):
        dev_count = st.number_input("Кол-во разработчиков", min_value=0)
        dev_status = f"Да ({dev_count} чел)"

# --- ЛОГИКА ОТЧЕТА ---
def make_expert_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Аудит ИТ и ИБ 2026"

    # Стили
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    bold = Font(bold=True)
    
    row = 1
    ws.merge_cells('A1:E1')
    ws['A1'] = f"ОТЧЕТ: {c_info.get('Наименование company', 'Аудит')}"
    ws['A1'].font = Font(size=14, bold=True, color="FFFFFF")
    ws['A1'].fill = header_fill
    row = 3

    # --- 1. АНАЛИЗ РИСКОВ (CTO/CISO View) ---
    ws.cell(row=row, column=1, value="КЛЮЧЕВЫЕ РИСКИ И РЕКОМЕНДАЦИИ").font = bold
    row += 1
    
    risks_found = []
    
    # Логика каналов
    ms = results.get("Интернет скорость", 0)
    bs = results.get("Резерв скорость", 0)
    if bs == 0:
        risks_found.append("КРИТИЧНО: Отсутствие резервного канала — единая точка отказа.")
    elif bs < (ms / 2):
        risks_found.append("ВНИМАНИЕ: Резервный канал не потянет нагрузку при отказе основного (менее 50% скорости).")

    # Логика Wi-Fi
    ap = results.get("WiFi Точки", 0)
    users = results.get("Пользователи", 0)
    if ap > 0:
        ratio = users / ap
        if ratio > 25:
            risks_found.append(f"ПЕРЕГРУЗКА WI-FI: На одну точку приходится {ratio:.1f} чел. Рекомендовано расширение.")
        if ap > 5 and results.get("WiFi Контроллер") == "Нет":
            risks_found.append("ПРОБЛЕМА УПРАВЛЕНИЯ: Более 5 точек без контроллера. Сложность настройки и риски безопасности.")

    # Логика масштабирования (SIEM / HelpDesk)
    total_hosts = results.get("Серверы", 0) + results.get("Виртуализация", 0) + users
    if total_hosts > 50 and results.get("SIEM") == "Нет":
        risks_found.append("CISO: Инфраструктура выросла. Необходим SIEM для мониторинга инцидентов (ELK, Positive Technologies, Splunk).")
    
    if users > 30 and results.get("HelpDesk") == "Нет":
        risks_found.append("CTO: Большой штат пользователей требует внедрения системы HelpDesk (Jira, GLPI, ITSM 365).")

    if results.get("Резервное копирование") == "Нет":
        risks_found.append("КРИТИЧНО: Данные не защищены. Требуется внедрение СРК уровня Enterprise.")

    if not risks_found:
        ws.cell(row=row, column=1, value="Критических замечаний по логике соответствия не найдено.")
        row += 1
    else:
        for r in risks_found:
            ws.cell(row=row, column=1, value=f"• {r}")
            row += 1
    
    row += 2
    # --- 2. ДЕТАЛЬНАЯ ТАБЛИЦА ---
    headers = ["Параметр", "Значение", "Статус", "Уровень рекомендации"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=row, column=i, value=h).font = white_font
        ws.cell(row=row, column=i).fill = header_fill
    row += 1

    for k, v in results.items():
        ws.cell(row=row, column=1, value=k)
        ws.cell(row=row, column=2, value=str(v))
        status = "OK" if "Нет" not in str(v) and v != 0 else "Требует внимания"
        ws.cell(row=row, column=3, value=status)
        ws.cell(row=row, column=4, value="CTO" if row % 2 == 0 else "CISO")
        row += 1

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['D'].width = 25
    
    wb.save(output)
    return output.getvalue()

# --- ФИНАЛИЗАЦИЯ ---
if st.button("📊 Сформировать экспертный отчет"):
    # Формируем итоговый словарь
    results = {
        "Интернет скорость": main_speed,
        "Резерв скорость": back_speed,
        "WiFi Точки": ap_cnt,
        "WiFi Контроллер": "Да" if wifi_ctrl else "Нет",
        "Пользователи": total_arm,
        "Маршрутизация": ", ".join(selected_routing) if selected_routing else "Нет",
        "NGFW": ngfw_vendor,
        "NAD (Сетевой анализ)": nad_v if nad else "Нет",
        "Серверы": phys_count,
        "Виртуализация": virt_count,
        "Резервное копирование": v_n_b,
        "СХД (Хранение)": storage_data,
        "HelpDesk": "Да" if helpdesk_exists else "Нет",
        "Web-ресурсы": web_status,
        "Разработка": dev_status
    }
    # Добавляем блок ИБ
    results.update(ib_values)

    with st.spinner("Генерация глубокой аналитики..."):
        report_bytes = make_expert_excel(client_info, results, score)
        
        # Telegram (silent)
        try:
            requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                          data={"chat_id": CHAT_ID, "caption": f"Audit: {client_info['Наименование компании']}"}, 
                          files={'document': (f"Audit_2026_{client_info['Наименование компании']}.xlsx", report_bytes)})
        except: pass

        st.success("Отчет сформирован с учетом корреляции данных!")
        st.download_button("📥 Скачать отчет (Excel)", report_bytes, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v9.1 | 2026 | Powered by AI Analytics")
