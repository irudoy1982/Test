import streamlit as st
import requests
import os
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="ИТ-Аудит 2026", layout="wide")

# Функция "Глубокой очистки" - удаляет вообще всё, кроме цифр и букв токена
def deep_clean(key):
    raw_value = st.secrets.get(key, "")
    # Удаляем кавычки, скобки, пробелы и невидимые символы
    clean_val = re.sub(r"[<>\s\"'“”‘’]", "", str(raw_value))
    return clean_val

TOKEN = deep_clean("TELEGRAM_TOKEN")
CHAT_ID = deep_clean("TELEGRAM_CHAT_ID")

# --- 2. ИНТЕРФЕЙС ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=200)
else:
    st.title("🛡️ Khalil Trade | IT Audit")

st.info(f"Текущий Chat ID в системе: `{CHAT_ID}`") # Выведем для проверки на экран

# Поля ввода
col1, col2 = st.columns(2)
with col1:
    company = st.text_input("Название компании", key="c_name")
    person = st.text_input("Имя", key="p_name")
with col2:
    email = st.text_input("Email", key="mail")
    phone = st.text_input("Телефон", key="phon")

# Данные для аудита
st.subheader("Технический опросник")
pc_count = st.number_input("Количество ПК", min_value=0)
has_backup = st.checkbox("Есть бэкап данных")

# --- 3. ГЕНЕРАЦИЯ EXCEL ---
def create_excel():
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit"
    ws['A1'] = "Параметр"; ws['B1'] = "Значение"
    ws['A2'] = "Компания"; ws['B2'] = company
    ws['A3'] = "ПК"; ws['B3'] = pc_count
    ws['A4'] = "Бэкап"; ws['B4'] = "Да" if has_backup else "Нет"
    wb.save(output)
    return output.getvalue()

# --- 4. ФУНКЦИЯ ОТПРАВКИ ---
def send_report(excel_file):
    # Прямая ссылка без лишних оберток
    api_url = f"https://api.telegram.org/bot{TOKEN}"
    
    text = (f"🚀 *НОВЫЙ АУДИТ*\n\n"
            f"🏢 Компания: {company}\n"
            f"👤 Контакт: {person}\n"
            f"📞 Тел: {phone}")
    
    try:
        # Отправляем текст
        r_msg = requests.post(
            f"{api_url}/sendMessage", 
            data={"chat_id": CHAT_ID, "text": text, "parse_mode": "Markdown"},
            timeout=10
        )
        
        if not r_msg.ok:
            return f"Ошибка текста: {r_msg.text}"

        # Отправляем файл
        files = {'document': (f"Audit_{company}.xlsx", excel_file)}
        r_file = requests.post(
            f"{api_url}/sendDocument", 
            data={"chat_id": CHAT_ID}, 
            files=files, 
            timeout=20
        )
        
        if r_file.ok:
            return True
        else:
            return f"Ошибка файла: {r_file.text}"
            
    except Exception as e:
        return f"Ошибка соединения: {str(e)}"

# --- 5. КНОПКА ЗАПУСКА ---
st.divider()
if st.button("🚀 СФОРМИРОВАТЬ И ОТПРАВИТЬ"):
    if not company or not person:
        st.warning("Заполните название компании и ваше имя!")
    elif not TOKEN or not CHAT_ID:
        st.error("Настройки TOKEN или CHAT_ID пусты в Secrets!")
    else:
        excel_raw = create_excel()
        with st.status("Связь с сервером Telegram...") as s:
            result = send_report(excel_raw)
            if result is True:
                s.update(label="✅ Отчет успешно доставлен!", state="complete")
                st.balloons()
            else:
                s.update(label="❌ Ошибка", state="error")
                st.error(result)
                st.write(f"Проверьте бота: https://t.me/{(TOKEN.split(':')[0] if ':' in TOKEN else '')}")

st.caption("Ivan Rudoy | 2026")
