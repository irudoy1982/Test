import streamlit as st
import requests
import os
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="ИТ-Аудит 2026", layout="wide", page_icon="🛡️")

# Глубокая очистка секретов
def deep_clean(key):
    raw_value = st.secrets.get(key, "")
    return re.sub(r"[<>\s\"'“”‘’]", "", str(raw_value))

TOKEN = deep_clean("TELEGRAM_TOKEN")
CHAT_ID = deep_clean("TELEGRAM_CHAT_ID")

# --- 2. ИНТЕРФЕЙС ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=250)
else:
    st.title("🛡️ Khalil Trade | IT Audit")

st.markdown("### Сбор данных для технического аудита")
st.divider()

# --- 3. ФОРМА ВВОДА (РЕШАЕТ ПРОБЛЕМУ ПЕРЕЗАГРУЗКИ) ---
with st.form("audit_form"):
    st.header("🏢 Основная информация")
    col1, col2 = st.columns(2)
    with col1:
        company_name = st.text_input("Наименование организации", placeholder="Введите название...")
        contact_person = st.text_input("ФИО контактного лица")
    with col2:
        email = st.text_input("Контактный email")
        phone = st.text_input("Контактный телефон")

    st.divider()
    st.header("📋 Технические параметры")
    
    c3, c4 = st.columns(2)
    with c3:
        pc_count = st.number_input("Количество АРМ (ПК)", min_value=0, step=1)
        srv_count = st.number_input("Количество серверов", min_value=0, step=1)
    with c4:
        backup = st.checkbox("Наличие резервного копирования")
        antivirus = st.checkbox("Наличие антивирусной защиты")
        ngfw = st.checkbox("Наличие межсетевого экрана (NGFW)")

    # Кнопка внутри формы
    submitted = st.form_submit_button("🚀 СФОРМИРОВАТЬ И ОТПРАВИТЬ ОТЧЕТ")

# --- 4. ЛОГИКА ГЕНЕРАЦИИ И ОТПРАВКИ ---
def create_excel(c_name, p_name, mail, tel, pc, srv, bkp, av, fw):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Report"
    
    # Стили
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    
    # Данные
    rows = [
        ["Параметр", "Значение"],
        ["Компания", c_name],
        ["Контакт", p_name],
        ["Email", mail],
        ["Телефон", tel],
        ["Кол-во ПК", pc],
        ["Кол-во серверов", srv],
        ["Резервное копирование", "Да" if bkp else "Нет"],
        ["Антивирус", "Да" if av else "Нет"],
        ["NGFW", "Да" if fw else "Нет"]
    ]
    
    for r_idx, row in enumerate(rows, 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            if r_idx == 1:
                cell.font = header_font
                cell.fill = header_fill
    
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 35
    wb.save(output)
    return output.getvalue()

def send_to_tg(c_name, p_name, score_text, excel_data):
    api_url = f"https://api.telegram.org/bot{TOKEN}"
    caption = (f"🔔 *НОВЫЙ АУДИТ*\n\n"
               f"🏢 Компания: {c_name}\n"
               f"👤 Контакт: {p_name}\n"
               f"📊 Параметры: {score_text}")
    
    try:
        # Отправляем сразу файл с подписью (это надежнее и быстрее)
        files = {'document': (f"Audit_{c_name}.xlsx", excel_data)}
        r = requests.post(
            f"{api_url}/sendDocument",
            data={"chat_id": CHAT_ID, "caption": caption, "parse_mode": "Markdown"},
            files=files,
            timeout=30
        )
        return True if r.ok else r.text
    except Exception as e:
        return str(e)

# --- 5. ОБРАБОТКА НАЖАТИЯ ---
if submitted:
    if not company_name or not contact_person:
        st.error("⚠️ Пожалуйста, заполните название компании и ФИО!")
    elif not TOKEN or not CHAT_ID:
        st.error("❌ Ошибка настройки: Проверьте Secrets (Token/ID)!")
    else:
        with st.spinner("Генерация и отправка..."):
            # Формируем краткий итог для текста
            summary = f"ПК: {pc_count}, Серверов: {srv_count}, Бэкап: {'✅' if backup else '❌'}"
            
            excel_raw = create_excel(company_name, contact_person, email, phone, 
                                     pc_count, srv_count, backup, antivirus, ngfw)
            
            result = send_to_tg(company_name, contact_person, summary, excel_raw)
            
            if result is True:
                st.success(f"✅ Отчет для '{company_name}' успешно отправлен в Telegram!")
                st.balloons()
                st.download_button("📥 Скачать копию Excel", excel_raw, f"Audit_{company_name}.xlsx")
            else:
                st.error(f"❌ Ошибка отправки: {result}")
                st.info("Проверьте, что в Secrets указан верный ID (число) и вы нажали START в боте.")

st.caption("Ivan Rudoy | IT Audit Pro 2026")
