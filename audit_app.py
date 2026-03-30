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

    1.  **Общая информация:** Укажите корректные контактные данные. Поле Email автоматически подстраивается под домен вашего сайта.
    2.  **Заполнение блоков:** Пройдите по разделам (ИТ, ИБ, Web, Разработка). Используйте переключатели (toggles) для активации нужных подразделов.
    3.  **Логический контроль (v6.2):** * Следите за тем, чтобы сумма устройств с разными ОС совпадала с общим числом АРМ.
        * При отсутствии критических решений ИБ система автоматически сформирует экспертное обоснование рисков.
    4.  **Результат:** После заполнения всех обязательных полей (*) нажмите кнопку «Сформировать экспертный отчет». Вы получите готовый файл Excel с анализом рисков и рекомендациями.
    """)

data = []
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
    country_codes = [("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"), ("🇰🇬 +996", "+996")]
    selected_code = p_col1.selectbox("Код", country_codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed")
    client_info['Контактный телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

st.divider()

# Вспомогательная функция для формирования данных (добавлено обоснование)
def add_audit_item(category, question, status, recommendation, is_risk=False):
    data.append({
        "Категория": category,
        "Объект": question,
        "Статус": status,
        "Рекомендация": recommendation,
        "is_risk": is_risk
    })

# --- БЛОК 1: ИТ-ИНФРАСТРУКТУРА ---
st.header("Блок 1: Информационные технологии")
col_it1, col_it2 = st.columns(2)
with col_it1:
    total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)
    selected_os = st.multiselect("ОС на АРМ:", ["Windows 10/11", "Linux", "macOS", "Legacy (XP/7)"])
    sum_os = 0
    for os_item in selected_os:
        val = st.number_input(f"Кол-во {os_item}:", min_value=0, step=1)
        sum_os += val
    if total_arm > 0 and sum_os != total_arm:
        st.warning(f"⚠️ Несовпадение: всего {total_arm}, указано {sum_os}")
        validation_errors.append("Ошибка баланса АРМ")

with col_it2:
    it_lic = st.radio("Используется ли лицензионное ПО?", ["Да", "Нет", "Частично"])
    if it_lic == "Да": score += 10; add_audit_item("ИТ", "Лицензионное ПО", "OK", "Соблюдается")
    elif it_lic == "Нет": add_audit_item("ИТ", "Лицензионное ПО", "РИСК", "Внедрить лицензионное ПО. Обоснование: минимизация юридических рисков и получение патчей безопасности.", True)
    else: score += 5; add_audit_item("ИТ", "Лицензионное ПО", "ВНИМАНИЕ", "Завершить лицензирование")

# --- БЛОК 2: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ (ДОБАВЛЕН MFA И ANTIDDOS) ---
st.divider()
st.header("Блок 2: Информационная безопасность")
col_ib1, col_ib2 = st.columns(2)

with col_ib1:
    # MFA Checkbox
    mfa_enabled = st.checkbox("Многофакторная аутентификация (MFA)", key="ib_mfa")
    if mfa_enabled:
        score += 20
        add_audit_item("ИБ", "MFA", "OK", "Внедрено")
    else:
        add_audit_item("ИБ", "MFA", "КРИТИЧЕСКИЙ РИСК", "Необходимо внедрить MFA. Обоснование: защита от компрометации учетных данных и несанкционированного доступа.", True)

    # Anti-DDoS Checkbox
    ddos_enabled = st.checkbox("Защита от DDoS-атак (Anti-DDoS)", key="ib_ddos")
    if ddos_enabled:
        score += 15
        add_audit_item("ИБ", "Anti-DDoS", "OK", "Внедрено")
    else:
        add_audit_item("ИБ", "Anti-DDoS", "РИСК", "Рекомендуется подключение Anti-DDoS. Обоснование: обеспечение доступности внешних сервисов при атаках на отказ в обслуживании.", True)

with col_ib2:
    av_enabled = st.checkbox("Централизованный Антивирус / EDR", key="ib_av")
    if av_enabled:
        score += 15
        add_audit_item("ИБ", "Антивирусная защита", "OK", "Внедрено")
    else:
        add_audit_item("ИБ", "Антивирусная защита", "КРИТИЧЕСКИЙ РИСК", "Требуется установка EDR. Обоснование: обнаружение и блокировка современных киберугроз на конечных точках.", True)

# --- 3. ГЕНЕРАЦИЯ EXCEL ---
def make_expert_excel(c_info, audit_data, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Результаты аудита"
    auto_date = datetime.now().strftime("%d.%m.%Y")
    
    # Стили
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Шапка отчета
    ws.append(["ОТЧЕТ ПО ТЕХНИЧЕСКОМУ АУДИТУ ИТ И ИБ (v6.2)"])
    ws.append([f"Компания: {c_info.get('Наименование компании')}", "", "", f"Дата: {auto_date}"])
    ws.append([f"Уровень защищенности: {final_score}%"])
    ws.append([])
    
    cols = ["Категория", "Объект аудита", "Статус", "Рекомендация и Обоснование"]
    ws.append(cols)
    for cell in ws[5]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border

    curr_row = 6
    for item in audit_data:
        is_risk = item.get("is_risk", False)
        ws.cell(row=curr_row, column=1, value=item["Категория"]).border = border
        ws.cell(row=curr_row, column=2, value=item["Объект"]).border = border
        
        st_cell = ws.cell(row=curr_row, column=3, value=item["Статус"])
        st_cell.border = border
        if is_risk: st_cell.font = Font(color="FF0000", bold=True)
        
        ws.cell(row=curr_row, column=4, value=item["Рекомендация"]).border = border
        ws.cell(row=curr_row, column=4).alignment = Alignment(wrap_text=True)
        curr_row += 1

    for col, width in {'A': 20, 'B': 30, 'C': 20, 'D': 65}.items():
        ws.column_dimensions[col].width = width
    
    wb.save(output)
    return output.getvalue(), auto_date

# --- ФИНАЛ ---
st.divider()
if validation_errors:
    st.error(f"🚨 Исправьте ошибки перед формированием отчета ({len(validation_errors)})")

if st.button("📊 Сформировать экспертный отчет", disabled=len(validation_errors) > 0):
    mandatory = [client_info['Город'], client_info['Наименование компании'], client_info['ФИО контактного лица'], client_info['Email']]
    if not all(mandatory):
        st.error("⚠️ Заполните все обязательные поля!")
    else:
        with st.spinner("Генерация отчета..."):
            f_score = min(score, 100)
            report_bytes, final_date = make_expert_excel(client_info, data, f_score)
            try:
                caption = f"🚀 *Новый аудит Khalil Trade*\n🏢 *{client_info['Наименование компании']}*\n📊 *Зрелость:* {f_score}%"
                requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                              data={"chat_id": CHAT_ID, "caption": caption, "parse_mode": "Markdown"},
                              files={"document": ("Audit_Report.xlsx", report_bytes)})
                st.success("✅ Отчет успешно сформирован и отправлен!")
                st.download_button("📥 Скачать Excel", report_bytes, f"Audit_{client_info['Наименование компании']}.xlsx")
            except Exception as e:
                st.error(f"Ошибка отправки: {e}")
