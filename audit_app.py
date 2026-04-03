import streamlit as st
import pandas as pd
import os
import requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Khalil Audit Expert v8.0", layout="wide", page_icon="🛡️")

# --- НАСТРОЙКИ TELEGRAM (из Secrets) ---
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ЛОГОТИП И КОНТАКТЫ ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### Экспертная система оценки рисков и финансового обоснования ИБ")
st.divider()

st.title("📋 Опросник: Технический аудит ИТ и ИБ v8.0 (Enterprise)")

# --- ИНСТРУКЦИЯ ДЛЯ ПОЛЬЗОВАТЕЛЯ ---
with st.expander("📖 Методология и инструкции v8.0"):
    st.markdown("""
    ### Новое в версии 8.0:
    1.  **Scale-Adaptive Logic:** Критичность отсутствия систем (SIEM, EDR, PAM) теперь зависит от масштаба вашей сети (кол-ва АРМ и серверов).
    2.  **Financial Risk Modeling:** Расчет потенциального ущерба от простоя на основе ваших бизнес-показателей.
    3.  **Vendor Best Practices:** Рекомендации адаптированы под стандарты NIST, CIS и ведущих вендоров (Cisco, Palo Alto, Microsoft).
    
    *Заполните общую информацию о компании, чтобы система смогла рассчитать финансовое обоснование рисков.*
    """)

data = {}
client_info = {}
validation_errors = []
score = 0

# --- ШАПКА: ИНФОРМАЦИЯ О КЛИЕНТЕ (Расширена для v8.0) ---
st.header("📍 Общая информация и Бизнес-контекст")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город:*")
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    client_info['Отрасль'] = st.selectbox("Отрасль:*", ["Производство", "Ритейл", "Финансы/Банки", "IT/Telecom", "Госсектор", "Энергетика", "Логистика", "Другое"])
    
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
    
    st.write("**Параметры для финансового обоснования:**")
    client_info['Revenue_Hour'] = st.number_input("Примерная выручка компании в час (KZT):", min_value=0, value=100000, step=10000, help="Используется для расчета потерь при простое (RTO)")
    client_info['RTO_Target'] = st.slider("Допустимое время восстановления (RTO), ч:", 1, 72, 4, help="Сколько часов бизнес может стоять без ИТ-систем")
    
    st.write("Контактный телефон:*")
    p_col1, p_col2 = st.columns([1, 2])
    country_codes = [("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"), ("🇦🇪 +971", "+971"), ("🇹🇷 +90", "+90")]
    selected_code = p_col1.selectbox("Код", country_codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed")
    client_info['Контактный телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

st.divider()

# --- БЛОК 1: ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ ---
st.header("Блок 1: Информационные технологии")

# 1.1 Конечные точки
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
    st.warning(f"⚠️ Ошибка в расчетах АРМ: Указано всего {total_arm}, а по ОС набралось {sum_os_arm}.")
    validation_errors.append("Несовпадение количества АРМ и ОС")

# 1.2 Сетевая инфраструктура
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_toggle"):
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL/VDSL", "Нет"]
    col_net1, col_net2 = st.columns(2)
    with col_net1:
        main_type = st.selectbox("Тип (основной):", net_types, key="main_net_type")
        main_speed = st.number_input("Скорость основного (Mbit/s):", min_value=0, step=10, key="main_net_speed")
        data['1.2.1. Основной канал'] = f"{main_type} ({main_speed} Mbit/s)"
    with col_net2:
        back_type = st.selectbox("Тип (резервный):", net_types, key="back_net_type")
        back_speed = st.number_input("Скорость резервного (Mbit/s):", min_value=0, step=10, key="back_net_speed")
        data['1.2.2. Резервный канал'] = f"{back_type} ({back_speed} Mbit/s)"

    if st.checkbox("Межсетевой экран (NGFW)", key="ngfw_chk"):
        ngfw_vendor = st.text_input("Производитель (NGFW):", key="ngfw_v")
        data['1.2.7. NGFW'] = f"Да ({ngfw_vendor if ngfw_vendor else 'не указан'})"
        score += 25
    else:
        data['1.2.7. NGFW'] = "Нет"

# 1.3 Серверы
st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
col_s1, col_s2 = st.columns(2)
with col_s1:
    phys_count = st.number_input("Количество физических серверов:", min_value=0, step=1, key="phys_srv")
    data['1.3.1. Физические серверы'] = phys_count
with col_s2:
    virt_count = st.number_input("Количество виртуальных серверов:", min_value=0, step=1, key="virt_srv")
    data['1.3.2. Виртуальные серверы'] = virt_count

if st.checkbox("Резервное копирование", key="ib_backup"):
    v_n_b = st.text_input("Вендор Резервного копирования:", key="vn_backup")
    data["Резервное копирование"] = f"Да ({v_n_b if v_n_b else 'не указан'})"
    score += 30
else:
    data["Резервное копирование"] = "Нет"

# 1.4 СХД
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Есть собственная СХД", key="storage_toggle"):
    st_media_sel = st.multiselect("Типы носителей:", ["HDD", "SSD", "NVMe"], key="st_media")
    data['1.4.1. Типы носителей'] = st_media_sel

# 1.5 Внутренние ИС
st.write("---")
st.subheader("1.5. Внутренние Информационные системы")
if st.toggle("ИС организации", key="is_toggle"):
    m_sys = st.selectbox("Почта:", ["Exchange", "Microsoft 365", "Google Workspace", "Собственный", "Нет"])
    data['1.5.1. Почтовая система'] = m_sys

st.divider()

# --- БЛОК 2: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ ---
st.header("Блок 2: Информационная Безопасность")
ib_systems = {
    "EPP (Антивирус)": 15, "DLP (Утечки)": 15, "PAM (Привилегии)": 10,
    "SIEM (Мониторинг)": 20, "VM (Уязвимости)": 10, "EDR/XDR (Точки)": 15, "WAF (Веб)": 10
}
col_ib1, col_ib2 = st.columns(2)
items = list(ib_systems.items())
for i, (label, pts) in enumerate(items):
    target_col = col_ib1 if i < 4 else col_ib2
    with target_col:
        if st.checkbox(label, key=f"ib_{label}"):
            v_n = st.text_input(f"Вендор {label}:", key=f"vn_{label}")
            data[label] = f"Да ({v_n if v_n else 'не указан'})"
            score += pts
        else:
            data[label] = "Нет"

st.divider()

# --- БЛОК 3: WEB-РЕСУРСЫ ---
st.header("Блок 3: Web-ресурсы")
if st.toggle("Web-ресурсы", key="web_toggle"):
    data['3.1. Хостинг'] = st.selectbox("Хостинг:", ["Собственный ЦОД", "Облако KZ", "Облако Global"])

st.divider()

# --- БЛОК 4: РАЗРАБОТКА ---
st.header("Блок 4: Разработка")
if st.toggle("Разработка", key="dev_toggle"):
    dev_count = st.number_input("Кол-во разработчиков:", min_value=0)
    data['4.1. Разработчики'] = dev_count
    data['4.2. CI/CD'] = st.checkbox("Используется CI/CD")

# --- ФУНКЦИЯ ГЕНЕРАЦИИ EXCEL v8.0 (Expert Logic) ---
def make_expert_excel_v8(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    
    # Стили
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    risk_fill = PatternFill(start_color="FF7C80", end_color="FF7C80", fill_type="solid")
    warning_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # ЛИСТ 1: EXECUTIVE SUMMARY (Для руководства)
    ws1 = wb.active
    ws1.title = "Executive Summary"
    
    ws1.merge_cells('A1:E2')
    ws1['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ ПО КИБЕРЗРЕЛОСТИ (v8.0 ALE)"
    ws1['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws1['A1'].font = Font(bold=True, size=16, color="1F4E78")

    # Расчет ALE (Annual Loss Expectancy - упрощенно как Impact per Incident)
    hourly_revenue = c_info.get('Revenue_Hour', 0)
    rto = c_info.get('RTO_Target', 4)
    incident_loss = hourly_revenue * rto
    
    ws1['A4'] = "БИЗНЕС-КОНТЕКСТ"; ws1['A4'].font = Font(bold=True, size=14)
    ws1['A5'] = "Компания:"; ws1['B5'] = c_info['Наименование компании']
    ws1['A6'] = "Отрасль:"; ws1['B6'] = c_info['Отрасль']
    ws1['A7'] = "Индекс зрелости:"; ws1['B7'] = f"{final_score}%"
    
    ws1['D5'] = "ФИНАНСОВЫЕ РИСКИ"; ws1['D5'].font = Font(bold=True, size=14)
    ws1['D6'] = "Стоимость часа простоя:"; ws1['E6'] = f"{hourly_revenue:,} KZT"
    ws1['D7'] = "Потери при инциденте (RTO):"; ws1['E7'] = f"{incident_loss:,} KZT"
    ws1['E7'].fill = risk_fill if final_score < 60 else warning_fill

    # ЛИСТ 2: ТЕХНИЧЕСКИЙ АУДИТ (Подробно)
    ws2 = wb.create_sheet("Technical Audit")
    headers = ["Параметр", "Значение", "Критичность (Scale-Adaptive)", "Финансовый риск", "Рекомендация"]
    for i, h in enumerate(headers, 1):
        cell = ws2.cell(row=1, column=i, value=h)
        cell.fill = header_fill; cell.font = white_font

    total_arm_val = results.get('1.1. Всего АРМ', 0)
    curr_row = 2
    for k, v in results.items():
        ws2.cell(row=curr_row, column=1, value=k).border = border
        ws2.cell(row=curr_row, column=2, value=str(v)).border = border
        
        # ЛОГИКА v8.0: Адаптивная критичность
        # Если компания крупная (>100 АРМ), отсутствие SIEM/NGFW/EDR становится КРИТИЧЕСКИМ
        is_missing = "Нет" in str(v) or v == 0
        criticality = "Средняя"
        if total_arm_val > 100:
            if any(x in k for x in ["SIEM", "EDR", "PAM", "NGFW", "Бэкап"]): criticality = "КРИТИЧЕСКАЯ"
        elif total_arm_val > 50:
            if any(x in k for x in ["NGFW", "Антивирус", "Бэкап"]): criticality = "ВЫСОКАЯ"
        
        if "XP" in str(k) or "2008" in str(k): criticality = "КРИТИЧЕСКАЯ (Legacy)"

        c_cell = ws2.cell(row=curr_row, column=3, value=criticality)
        c_cell.border = border
        if criticality == "КРИТИЧЕСКАЯ" and is_missing: c_cell.fill = risk_fill

        # Расчет доли финансового риска
        fin_risk_val = "N/A"
        if is_missing:
            risk_factor = 0.8 if criticality == "КРИТИЧЕСКАЯ" else 0.4 if criticality == "ВЫСОКАЯ" else 0.1
            fin_risk_val = f"~{int(incident_loss * risk_factor):,} KZT"
        
        ws2.cell(row=curr_row, column=4, value=fin_risk_val).border = border
        
        # Рекомендация на основе бест-практис
        rec = "Поддерживать состояние"
        if is_missing:
            if "SIEM" in k: rec = "Внедрение SOC/SIEM для мониторинга атак"
            elif "NGFW" in k: rec = "Установка межсетевого экрана нового поколения"
            elif "Бэкап" in k: rec = "Настройка резервного копирования по правилу 3-2-1"
            else: rec = "Рассмотреть внедрение в рамках бюджета"
        
        ws2.cell(row=curr_row, column=5, value=rec).border = border
        curr_row += 1

    # Настройка ширины колонок
    for sheet in [ws1, ws2]:
        for col in sheet.columns:
            sheet.column_dimensions[col[0].column_letter].width = 30

    wb.save(output)
    return output.getvalue()

# --- ФИНАЛЬНАЯ ЧАСТЬ ---
st.divider()

if validation_errors:
    st.error(f"🚨 Исправьте ошибки в данных. Проблем: {len(validation_errors)}")

if st.button("📊 Сформировать экспертный отчет v8.0", disabled=len(validation_errors) > 0):
    mandatory = [client_info['Город'], client_info['Наименование компании'], client_info['ФИО контактного лица'], client_info['Email']]
    if not all(mandatory):
        st.error("⚠️ Пожалуйста, заполните все обязательные поля (*)")
    else:
        with st.spinner("Проводим анализ рисков и расчет финансового ущерба..."):
            f_score = min(score, 100)
            report_bytes = make_expert_excel_v8(client_info, data, f_score)
            
            # Telegram Notification
            try:
                caption = (f"🚀 *Новый аудит Khalil v8.0*\n"
                           f"🏢 *{client_info['Наименование компании']}*\n"
                           f"📊 *Зрелость:* {f_score}%\n"
                           f"💰 *Потери (RTO):* {client_info['Revenue_Hour']*client_info['RTO_Target']:,} KZT")
                requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                              data={"chat_id": CHAT_ID, "caption": caption, "parse_mode": "Markdown"}, 
                              files={'document': (f"Audit_v8_{client_info['Наименование компании']}.xlsx", report_bytes)})
            except: pass
            
            st.success("Экспертный отчет v8.0 готов!")
            st.download_button(
                label="📥 Скачать отчет (Excel)",
                data=report_bytes,
                file_name=f"Khalil_Audit_v8_{client_info['Наименование компании']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

st.info("Khalil Audit System v8.0 | Almaty 2026 | Risk-Based Approach")
