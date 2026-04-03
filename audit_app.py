import streamlit as st
import pandas as pd
import os
import requests
import re
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

st.title("📋 Опросник Технический аудит ИТ и ИБ (2026) v7.0b")

# --- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ API ---
def get_eol_status(product_key, raw_version):
    """
    Очищает версию через RegEx и проверяет статус EoL через API endoflife.date
    """
    if not raw_version or not product_key:
        return None

    # Очистка: оставляем только цифры и точки
    clean_version = re.sub(r'[^0-9.]', '', str(raw_version)).strip('.')
    
    # Извлекаем основной цикл (н-р, 2013 или 11)
    version_parts = clean_version.split('.')
    cycle = version_parts[0] if version_parts else None

    if not cycle:
        return None

    url = f"https://endoflife.date/api/{product_key}/{cycle}.json"
    
    try:
        response = requests.get(url, timeout=2)
        if response.status_code == 200:
            data = response.json()
            eol_value = data.get('eol', False)
            
            if isinstance(eol_value, str):
                eol_date = datetime.strptime(eol_value, '%Y-%m-%d')
                if eol_date < datetime.now():
                    return f"EoL: Снято с поддержки ({eol_value})"
                return f"Поддерживается до {eol_value}"
            elif eol_value is True:
                return "EoL: Снято с поддержки"
        return "Статус актуален или версия не найдена"
    except:
        return "Сервис проверки недоступен"

# --- ИНСТРУКЦИЯ ДЛЯ ПОЛЬЗОВАТЕЛЯ ---
with st.expander("📖 Инструкция по заполнению (нажмите, чтобы развернуть)"):
    st.markdown("""
    ### Руководство по проведению экспресс-аудита
    
    Данный инструмент предназначен для сбора технических данных об ИТ-ландшафте и уровне защищенности вашей организации. Пожалуйста, следуйте шагам ниже:

    1. **Общая информация:** Укажите корректные контактные данные. Все поля со звездочкой (*) обязательны.
    2. **Заполнение блоков:** Пройдите по разделам. Используйте переключатели (toggles) для активации нужных подразделов.
    3. **Логический контроль:** Сумма ОС на АРМ должна быть равна общему числу АРМ. Количество ОС на серверах должно быть не меньше числа вирт. машин.
    4. **Результат:** Нажмите кнопку «Сформировать экспертный отчет» для получения файла Excel с автоматической проверкой версий ПО.
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
    custom_email_mode = st.checkbox("Email отличается от сайта")
    
    if custom_email_mode:
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

if not all([client_info.get('Город'), client_info.get('Наименование компании'), client_info.get('Сайт компании'), client_info.get('Email'), client_info.get('ФИО контактного лица'), client_info.get('Должность'), phone_num]):
    validation_errors.append("Заполните все обязательные поля в блоке 'Общая информация'")

st.divider()

# --- БЛОК 1: ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ ---
st.header("Блок 1: Информационные технологии")

st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт)", min_value=0, step=1)
data['1.1. Всего АРМ'] = total_arm
selected_os_arm = st.multiselect("ОС на АРМ", ["Windows XP/Vista/7/8", "Windows 10", "Windows 11", "Linux", "macOS", "Другое"])

sum_os_arm = 0
if selected_os_arm:
    for os_item in selected_os_arm:
        val = st.number_input(f"Кол-во на {os_item}", min_value=0, step=1, key=f"arm_{os_item}")
        data[f"ОС АРМ ({os_item})"] = val
        sum_os_arm += val
data['1.1. Примечание'] = st.text_area("Примечание к разделу 1.1", key="note_1_1")

if total_arm > 0 and sum_os_arm != total_arm:
    validation_errors.append("Несовпадение количества АРМ и ОС")

st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_toggle"):
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL/VDSL", "Нет"]
    routing_types = ["Статическая", "RIP", "OSPF", "EIGRP", "BGP", "IS-IS"]
    col_net1, col_net2 = st.columns(2)
    with col_net1:
        main_type = st.selectbox("Тип (основной)", net_types, key="main_net_type")
        main_speed = st.number_input("Скорость основного (Mbit/s)", min_value=0, key="main_net_speed")
        data['1.2.1. Основной канал'] = f"{main_type} ({main_speed} Mbit/s)"
    with col_net2:
        back_type = st.selectbox("Тип (резервный)", net_types, index=6, key="back_net_type")
        back_speed = st.number_input("Скорость резервного (Mbit/s)", min_value=0, key="back_net_speed")
        data['1.2.2. Резервный канал'] = f"{back_type} ({back_speed} Mbit/s)"

    selected_routing = st.multiselect("Тип маршрутизации*", routing_types, key="routing_sel")
    data['1.2.3. Маршрутизация'] = ", ".join(selected_routing)
    
    if st.checkbox("Межсетевой экран (NGFW)", key="ngfw_chk"):
        ngfw_vendor = st.text_input("Производитель (NGFW)", key="ngfw_v")
        data['1.2.7. NGFW'] = f"Да ({ngfw_vendor})" if ngfw_vendor else "Нет"
        score += 20

st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
phys_count = st.number_input("Количество физических серверов", min_value=0, step=1, key="phys_srv")
data['1.3.1. Физические серверы'] = phys_count
virt_count = st.number_input("Количество виртуальных серверов", min_value=0, step=1, key="virt_srv")
data['1.3.2. Виртуальные серверы'] = virt_count

s_os_list = ["Windows Server 2008/2012 R2", "Windows Server 2016", "Windows Server 2019", "Windows Server 2022", "Linux", "Unix", "Другое"]
selected_os_srv = st.multiselect("Выберите ОС серверов", s_os_list, key="ms_srv_list")
if selected_os_srv:
    for os_s in selected_os_srv:
        val_os = st.number_input(f"Кол-во на {os_s}", min_value=0, key=f"fsrv_{os_s}")
        data[f"ОС Сервера ({os_s})"] = val_os

if st.checkbox("Резервное копирование", key="ib_backup"):
    v_n_b = st.text_input("Вендор Резервного копирования", key="vn_backup")
    data["Резервное копирование"] = v_n_b
    score += 20

st.write("---")
st.subheader("1.5. Внутренние Информационные системы")
if st.toggle("ИС организации", key="is_toggle"):
    m_opts = ["Exchange (On-Prem)", "Lotus", "Microsoft 365", "Google Workspace", "Собственный", "Нет"]
    m_sys = st.selectbox("Почтовая система", m_opts)
    if m_sys == "Exchange (On-Prem)":
        m_ver = st.text_input("Версия Exchange (н-р: 2013, 2019)*", key="mail_version_input")
        data['1.5.1. Почтовая система'] = f"Exchange (v.{m_ver})"
    else:
        data['1.5.1. Почтовая система'] = m_sys

st.divider()

# --- БЛОК 2: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ ---
st.header("Блок 2: Информационная Безопасность")
if st.toggle("Средства защиты", key="ib_toggle"):
    ib_systems = {"EPP (Антивирус)": 10, "DLP (Утечки)": 15, "SIEM (Мониторинг)": 20}
    for label, pts in ib_systems.items():
        if st.checkbox(label, key=f"fib_{label}"):
            v_n = st.text_input(f"Вендор {label}*", key=f"fvn_{label}")
            data[label] = f"Да ({v_n})"
            score += pts
        else:
            data[label] = "Нет"

st.divider()

# --- БЛОК 3: WEB И БЛОК 4: РАЗРАБОТКА (Упрощено для краткости) ---
st.header("Блоки 3 и 4: Web и Разработка")
if st.toggle("Показать блоки", key="dev_web_toggle"):
    data['4.2. CICD'] = st.checkbox("Используется CI/CD", key="cicd_f")

# --- ГЕНЕРАЦИЯ EXCEL ---
def make_expert_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Khalil Audit Report"
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    ws.merge_cells('A1:D2')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ ПО ИТ И ИБ (2026)"
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=16, color="1F4E78")
    
    curr_row = 4
    for k, v in c_info.items():
        ws.cell(row=curr_row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=curr_row, column=2, value=str(v))
        curr_row += 1
    
    auto_date = datetime.now().strftime("%d.%m.%Y %H:%M")
    ws.cell(row=curr_row, column=1, value="ИНДЕКС ЗРЕЛОСТИ").font = Font(bold=True)
    score_cell = ws.cell(row=curr_row, column=2, value=f"{final_score}%")
    bg = "92D050" if final_score > 70 else "FFC000" if final_score > 40 else "FF7C80"
    score_cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
    
    curr_row += 3
    headers = ["Параметр", "Значение", "Статус", "Рекомендация (включая API Check)"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=curr_row, column=i, value=h)
        cell.fill = header_fill; cell.font = white_font
    
    curr_row += 1
    for k, v in results.items():
        val_str = str(v).lower()
        status, rec = "В норме", "Поддерживать состояние."
        is_risk = False

        # --- ЛОГИКА ПРОВЕРКИ ЧЕРЕЗ API ---
        api_info = None
        if "Exchange" in str(v):
            ver_match = re.search(r'v\.?([\d.]+)', val_str)
            if ver_match: api_info = get_eol_status("exchange", ver_match.group(1))
        elif "Windows" in k:
            # Для Windows 10/11/Server
            ver_match = re.search(r'(\d+)', k)
            if ver_match: api_info = get_eol_status("windows", ver_match.group(1))

        if api_info:
            rec = f"API Check: {api_info}."
            if "Снято" in api_info: status = "КРИТИЧНО"; is_risk = True

        # --- СТАНДАРТНАЯ ЛОГИКА ---
        if "Примечание" in k: 
            status, rec = "Инфо", "Учтено."
        elif any(old in k for old in ["XP", "7", "8", "2008", "2012"]) and isinstance(v, (int, float)) and v > 0:
            status, rec = "КРИТИЧНО", "Система устарела (EoL). Срочная замена."
            is_risk = True
        elif "Нет" in str(v) or v == 0 or not val_str.strip():
            if status == "В норме":
                is_risk = True; status = "РИСК"; rec = "Рассмотреть внедрение."

        ws.cell(row=curr_row, column=1, value=k).border = border
        ws.cell(row=curr_row, column=2, value=str(v)).border = border
        st_cell = ws.cell(row=curr_row, column=3, value=status)
        st_cell.border = border
        if is_risk: st_cell.font = Font(color="FF0000", bold=True)
        
        ws.cell(row=curr_row, column=4, value=rec).border = border
        ws.cell(row=curr_row, column=4).alignment = Alignment(wrapText=True)
        curr_row += 1
        
    for col, width in {'A': 35, 'B': 30, 'C': 15, 'D': 55}.items():
        ws.column_dimensions[col].width = width
    wb.save(output)
    return output.getvalue(), auto_date

# --- ФИНАЛ ---
st.divider()
if validation_errors:
    st.error(f"🚨 Ошибок: {len(validation_errors)}")
    for err in validation_errors: st.write(f"- {err}")

if st.button("📊 Сформировать экспертный отчет", disabled=len(validation_errors) > 0):
    with st.spinner("Проверка версий через API и генерация отчета..."):
        f_score = min(score, 100)
        report_bytes, _ = make_expert_excel(client_info, data, f_score)
        st.success("Отчет готов!")
        st.download_button("📥 Скачать отчет", report_bytes, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v7.1 (API Enabled) | Almaty 2026")
