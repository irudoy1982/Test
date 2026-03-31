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

# --- 2. ЛОГОТИП И КОНТАКТЫ ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### Мы поможем Вам стать лучше!")
st.divider()

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v6.2")

# --- ИНСТРУКЦИЯ ДЛЯ ПОЛЬЗОВАТЕЛЯ ---
with st.expander("📖 Инструкция по заполнению (нажмите, чтобы развернуть)"):
    st.markdown("""
    ### Руководство по проведению экспресс-аудита
    
    Данный инструмент предназначен для сбора технических данных об ИТ-ландшафте и уровне защищенности вашей организации. Пожалуйста, следуйте шагам ниже:

    1.  **Общая информация:** Укажите корректные контактные данные.
    2.  **Заполнение блоков:** Пройдите по разделам (ИТ, ИБ, Web, Разработка).
    3.  **Логический контроль:** Система проверяет соответствие количества ОС общему числу устройств.
    4.  **Результат:** После заполнения нажмите «Сформировать экспертный отчет». Умная система проанализирует риски на основе масштаба вашей сети.
    """)

data = {}
client_info = {}
validation_errors = []
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
        client_info['Email'] = st.text_input("Email контактного лица:*")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица (логин):*")
            e_col1, e_col2 = st.columns([2, 3])
            with e_col1:
                email_prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="email_pre")
            with e_col2:
                st.markdown(f"<div style='padding-top: 5px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
            client_info['Email'] = f"{email_prefix}@{clean_domain}" if email_prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    st.write("Контактный телефон:*")
    p_col1, p_col2 = st.columns([1, 2])
    country_codes = [("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998")]
    selected_code = p_col1.selectbox("Код", country_codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed")
    client_info['Контактный телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

st.divider()

# --- БЛОК 1: ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ ---
st.header("Блок 1: Информационные технологии")
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)
data['1.1. Всего АРМ'] = total_arm
selected_os_arm = st.multiselect("ОС на АРМ:", ["Windows XP/Vista/7/8", "Windows 10", "Windows 11", "Linux", "macOS", "Другое"])
sum_os_arm = 0
if selected_os_arm:
    for os_item in selected_os_arm:
        val = st.number_input(f"Кол-во на {os_item}:", min_value=0, step=1, key=f"arm_{os_item}")
        data[f"ОС АРМ ({os_item})"] = val
        sum_os_arm += val
if total_arm > 0 and sum_os_arm != total_arm:
    st.warning(f"⚠️ Ошибка в расчетах АРМ: Указано {total_arm}, по ОС {sum_os_arm}.")
    validation_errors.append("Несовпадение АРМ")

st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_toggle"):
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL/VDSL", "Нет"]
    c1, c2 = st.columns(2)
    with c1:
        main_type = st.selectbox("Тип (основной):", net_types)
        main_speed = st.number_input("Скорость (Mbit/s):", min_value=0)
        data['1.2.1. Основной канал'] = f"{main_type} ({main_speed})"
    with c2:
        if st.checkbox("Межсетевой экран (NGFW)"):
            ngfw_v = st.text_input("Вендор NGFW:")
            data['1.2.7. NGFW'] = f"Да ({ngfw_v})"
            score += 20

st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
col_s1, col_s2 = st.columns(2)
with col_s1:
    phys_count = st.number_input("Физические серверы:", min_value=0, step=1)
    data['1.3.1. Физические серверы'] = phys_count
with col_s2:
    virt_count = st.number_input("Виртуальные серверы:", min_value=0, step=1)
    data['1.3.2. Виртуальные серверы'] = virt_count

s_os_list = ["Windows Server 2008/2012 R2", "Windows Server 2016", "Windows Server 2019", "Windows Server 2022", "Linux", "Unix", "Другое"]
selected_os_srv = st.multiselect("ОС серверов:", s_os_list)
sum_os_srv = 0
for os_s in selected_os_srv:
    val_os = st.number_input(f"Кол-во на {os_s}:", min_value=0)
    data[f"ОС Сервера ({os_s})"] = val_os
    sum_os_srv += val_os

if st.checkbox("Резервное копирование"):
    v_n_b = st.text_input("Вендор СРК:")
    data["Резервное копирование"] = f"Да ({v_n_b})"
    score += 20

st.write("---")
st.subheader("1.5. Внутренние Информационные системы")
is_active = st.toggle("ИС организации")
if is_active:
    m_sys = st.selectbox("Почта:", ["Exchange (On-Prem)", "Lotus", "Microsoft 365", "Google Workspace", "Собственный", "Нет"])
    if m_sys in ["Exchange (On-Prem)", "Lotus"]:
        m_ver = st.text_input(f"Версия {m_sys}:")
        data['1.5.1. Почтовая система'] = f"{m_sys} (v.{m_ver})"
    else:
        data['1.5.1. Почтовая система'] = m_sys

    for label in ["1С", "Битрикс24", "Documentolog"]:
        if st.checkbox(label): data[f"ИС: {label}"] = st.text_input(f"Версия {label}:")

st.divider()

# --- БЛОК 2: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ ---
st.header("Блок 2: Информационная Безопасность")
ib_active = st.toggle("Средства защиты", key="ib_toggle")
ib_results = {} # Для временного хранения состояния чекбоксов
if ib_active:
    ib_systems = {
        "EPP (Антивирус)": 10, "DLP (Утечки)": 15, "PAM (Привилегии)": 10,
        "SIEM (Мониторинг)": 20, "VM (Уязвимости)": 10, "EDR/XDR (Точки)": 15,
        "WAF (Веб)": 10, "Sandbox (Песочница)": 5, "IDS/IPS (Атаки)": 5, "IDM/IGA (Доступ)": 5,
        "MFA (Аутентификация)": 15, "Anti-DDoS": 15
    }
    col_ib1, col_ib2 = st.columns(2)
    items = list(ib_systems.items())
    for i, (label, pts) in enumerate(items):
        target_col = col_ib1 if i < 6 else col_ib2
        with target_col:
            if st.checkbox(label):
                v_n = st.text_input(f"Вендор {label}:", key=f"vn_{label}")
                data[label] = f"Да ({v_n if v_n else 'не указан'})"
                ib_results[label] = True
                score += pts
            else:
                data[label] = "Нет"
                ib_results[label] = False

st.divider()

# --- БЛОК 3: WEB-РЕСУРСЫ ---
st.header("Блок 3: Web-ресурсы")
web_active = st.toggle("Web-ресурсы", key="web_toggle")
if web_active:
    data['3.1. Хостинг'] = st.selectbox("Хостинг:", ["Собственный ЦОД", "Облако KZ", "Облако Global"])
    data['3.2. Frontend'] = st.multiselect("Frontend:", ["Nginx", "Apache", "IIS"])

# --- БЛОК 4: РАЗРАБОТКА ---
st.header("Блок 4: Разработка")
dev_active = st.toggle("Разработка", key="dev_toggle")
if dev_active:
    data['4.1. Разработчики'] = st.number_input("Кол-во разработчиков:", min_value=0)

# --- ГЕНЕРАЦИЯ EXCEL (С УМНОЙ ЛОГИКОЙ) ---
def make_expert_excel(c_info, results, final_score, app_data):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Khalil Audit Report"
    
    # Стили
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Шапка
    ws.merge_cells('A1:D2')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ ПО ИТ И ИБ (2026)"
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=14)

    # Информация о клиенте
    curr_row = 4
    for k, v in c_info.items():
        ws.cell(row=curr_row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=curr_row, column=2, value=str(v))
        curr_row += 1

    # Индекс зрелости
    ws.cell(row=curr_row, column=1, value="ИНДЕКС ЗРЕЛОСТИ:").font = Font(bold=True)
    score_cell = ws.cell(row=curr_row, column=2, value=f"{final_score}%")
    curr_row += 3

    # Заголовки таблицы
    headers = ["Параметр", "Значение", "Статус", "Рекомендация / Обоснование"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=curr_row, column=i, value=h)
        cell.fill = header_fill; cell.font = white_font; cell.border = border

    curr_row += 1
    
    # Данные для умной логики
    n_arm = app_data.get('1.1. Всего АРМ', 0)
    n_srv = app_data.get('1.3.1. Физические серверы', 0) + app_data.get('1.3.2. Виртуальные серверы', 0)
    has_web = web_active or dev_active
    
    for k, v in results.items():
        ws.cell(row=curr_row, column=1, value=k).border = border
        ws.cell(row=curr_row, column=2, value=str(v)).border = border
        
        status = "В норме"
        rec = "Риски минимизированы."
        font_color = "000000"

        # УМНАЯ ЛОГИКА ДЛЯ ПРОДУКТОВ ИБ
        if v == "Нет":
            if k == "PAM (Привилегии)":
                if n_srv < 10:
                    status, rec = "ВНИМАНИЕ", "Малое кол-во серверов. Рассмотреть в будущем при росте парка."
                else:
                    status, rec = "КРИТИЧНО", "Высокий риск компрометации админ-аккаунтов при 10+ серверах."
            
            elif k == "SIEM (Мониторинг)":
                if n_arm < 200 and n_srv < 20 and not is_active:
                    status, rec = "ВНИМАНИЕ", "Инфраструктура невелика. Достаточно ручного анализа логов."
                else:
                    status, rec = "КРИТИЧНО", "Необходим сбор событий для обнаружения атак в реальном времени."
            
            elif k == "VM (Уязвимости)":
                if n_arm < 100 and n_srv < 10:
                    status, rec = "ВНИМАНИЕ", "Рассмотреть к приобретению по мере усложнения сети."
                else:
                    status, rec = "КРИТИЧНО", "Высокий риск эксплуатации старых уязвимостей в крупной сети."
            
            elif k == "EDR/XDR (Точки)":
                if n_arm < 50:
                    status, rec = "РЕКОМЕНДУЕТСЯ", "Достаточно классического EPP при малом кол-ве АРМ."
                else:
                    status, rec = "КРИТИЧНО", "Необходим расширенный анализ угроз на конечных точках."
            
            elif k == "WAF (Веб)":
                if not has_web:
                    status, rec = "НЕ ТРЕБУЕТСЯ", "Отсутствуют публичные веб-сервисы и разработка."
                else:
                    status, rec = "КРИТИЧНО", "Веб-ресурсы без защиты подвержены атакам SQLi, XSS."
            
            elif k in ["IDM/IGA (Доступ)", "Anti-DDoS"]:
                status, rec = "ВАЖНО", "Рекомендуется рассмотреть к приобретению для масштабируемости и устойчивости."
            
            elif k == "MFA (Аутентификация)":
                if n_arm < 20:
                    status, rec = "НИЗКИЙ РИСК", "При малом штате достаточно строгих парольных политик."
                else:
                    status, rec = "КРИТИЧНО", "Двухфакторная защита обязательна для предотвращения кражи данных."
            
            else:
                # Общая логика для остальных (DLP, Sandbox и т.д.)
                status, rec = "КРИТИЧНО", f"Отсутствие {k} повышает риск реализации ИБ-угроз."

            if status == "КРИТИЧНО": font_color = "FF0000"
            elif status in ["ВНИМАНИЕ", "ВАЖНО"]: font_color = "FFC000"

        st_cell = ws.cell(row=curr_row, column=3, value=status)
        st_cell.font = Font(color=font_color, bold=True)
        st_cell.border = border
        ws.cell(row=curr_row, column=4, value=rec).border = border
        curr_row += 1

    for col, width in {'A': 30, 'B': 25, 'C': 15, 'D': 50}.items():
        ws.column_dimensions[col].width = width
    
    wb.save(output)
    return output.getvalue()

# --- ФИНАЛ ---
st.divider()
if validation_errors:
    st.error(f"🚨 Исправьте ошибки ({len(validation_errors)}) перед выгрузкой.")

if st.button("📊 Сформировать экспертный отчет", disabled=len(validation_errors) > 0):
    if not all([client_info['Город'], client_info['Наименование компании'], client_info['Email']]):
        st.error("⚠️ Заполните обязательные поля!")
    else:
        with st.spinner("Анализируем риски..."):
            f_score = min(score, 100)
            report_bytes = make_expert_excel(client_info, data, f_score, data)
            try:
                cap = f"🚀 *Аудит:* {client_info['Наименование компании']}\n📊 *Зрелость:* {f_score}%"
                requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                              data={"chat_id": CHAT_ID, "caption": cap, "parse_mode": "Markdown"}, 
                              files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report_bytes)})
            except: pass
            st.success("Отчет сформирован с учетом масштаба вашей сети!")
            st.download_button("📥 Скачать Excel", report_bytes, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v6.2 | Smart Logic Enabled")
