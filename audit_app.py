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
    
    Данный инструмент предназначен для сбора технических данных об ИТ-ландшафте и уровне защищенности вашей организации.

    1.  **Общая информация:** Укажите корректные контактные данные.
    2.  **Заполнение блоков:** Пройдите по разделам (ИТ, ИБ).
    3.  **Логический контроль (v6.2):** * Следите за совпадением количества АРМ и указанных ОС.
        * При отсутствии критических решений ИБ система автоматически сформирует обоснование для руководства.
    4.  **Результат:** После заполнения нажмите кнопку «Сформировать экспертный отчет». Вы получите файл Excel с анализом рисков и рекомендациями.
    """)

data_results = []
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

# Вспомогательная функция для вопросов
def audit_question(category, question, weight, justification):
    global score
    st.write(f"**{question}**")
    ans = st.radio(f"Выбор для: {question}", ["Да", "Нет", "В процессе"], key=question, label_visibility="collapsed", horizontal=True)
    
    status = "OK"
    rec = ""
    is_risk = False
    
    if ans == "Да":
        score += weight
    elif ans == "Нет":
        status = "КРИТИЧЕСКИЙ РИСК"
        rec = f"Рекомендуется внедрение решения. Обоснование: {justification}"
        is_risk = True
    else:
        score += (weight / 2)
        status = "ВНИМАНИЕ"
        rec = "Требуется завершить внедрение и провести аудит настроек."
        
    data_results.append({
        "Категория": category,
        "Объект": question,
        "Статус": status,
        "Рекомендация": rec,
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
    audit_question("ИТ", "Используется ли лицензионное ПО на критических узлах?", 10, "Использование нелицензионного ПО ведет к отсутствию обновлений безопасности и юридическим рискам.")
    audit_question("ИТ", "Проводится ли регулярное резервное копирование?", 15, "Без актуальных копий восстановление бизнеса после сбоя или атаки невозможно.")

# --- БЛОК 2: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ (ОБНОВЛЕННЫЙ) ---
st.divider()
st.header("Блок 2: Информационная безопасность")
col_ib1, col_ib2 = st.columns(2)

with col_ib1:
    audit_question("ИБ", "Внедрена ли многофакторная аутентификация (MFA)?", 20, 
                   "MFA предотвращает 99% атак на учетные данные, защищая доступ даже в случае кражи пароля.")
    
    audit_question("ИБ", "Используется ли система PAM для контроля администраторов?", 15, 
                   "Системы Privileged Access Management минимизируют риск злоупотребления высокими правами доступа и утечки данных через привилегированных пользователей.")

with col_ib2:
    audit_question("ИБ", "Используется ли защита рабочих станций класса EDR/XDR?", 15, 
                   "Обычные антивирусы не справляются с современными шифровальщиками. EDR позволяет обнаружить атаку на ранней стадии.")
    
    audit_question("ИБ", "Внедрены ли неизменяемые (Immutable) бэкапы?", 10, 
                   "Современные вирусы-шифровальщики сначала удаляют бэкапы. Неизменяемые хранилища делают удаление невозможным.")

# --- 3. ГЕНЕРАЦИЯ EXCEL ---
def make_expert_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Аудит ИТ и ИБ 2026"
    
    # Стили
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Шапка
    ws.append(["ОТЧЕТ ПО ТЕХНИЧЕСКОМУ АУДИТУ (v6.2)"])
    ws.append([f"Компания: {c_info.get('Наименование компании')}"])
    ws.append([f"Уровень зрелости: {final_score}%"])
    ws.append([])
    
    headers = ["Категория", "Объект аудита", "Текущий статус", "Экспертная рекомендация и Обоснование"]
    ws.append(headers)
    for cell in ws[5]:
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border

    # Данные
    for i, item in enumerate(results, 6):
        is_risk = item.pop("is_risk", False)
        row_data = list(item.values())
        ws.append(row_data)
        
        # Оформление ячеек
        for col in range(1, 5):
            cell = ws.cell(row=i, column=col)
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical='top')
            if is_risk and col == 3:
                cell.font = Font(color="FF0000", bold=True)

    # Ширина колонок
    for col, width in {'A': 15, 'B': 35, 'C': 20, 'D': 65}.items():
        ws.column_dimensions[col].width = width
    
    wb.save(output)
    return output.getvalue()

# --- ФИНАЛ ---
st.divider()
if st.button("📊 Сформировать экспертный отчет", disabled=len(validation_errors) > 0):
    if not all([client_info['Город'], client_info['Наименование компании'], client_info['Email']]):
        st.error("⚠️ Заполните обязательные поля (Город, Компания, Email)!")
    else:
        with st.spinner("Формируем отчет..."):
            f_score = min(score, 100)
            report_bytes = make_expert_excel(client_info, data_results, f_score)
            
            try:
                # Отправка в Telegram
                caption = f"🛡️ *Новый аудит (v6.2)*\n🏢 {client_info['Наименование компании']}\n📊 Зрелость: {f_score}%"
                requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                              data={"chat_id": CHAT_ID, "caption": caption, "parse_mode": "Markdown"},
                              files={"document": ("Audit_Report.xlsx", report_bytes)})
                
                st.success("✅ Отчет сформирован!")
                st.download_button("📥 Скачать отчет (Excel)", report_bytes, "Audit_2026.xlsx")
            except Exception as e:
                st.error(f"Ошибка отправки: {e}")
