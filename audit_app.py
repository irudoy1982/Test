import streamlit as st
import requests
import re
from io import BytesIO
from openpyxl import Workbook

# --- 1. ОЧИСТКА ДАННЫХ ---
def clean(key):
    val = st.secrets.get(key, "")
    return re.sub(r"[^a-zA-Z0-9:]", "", str(val)) if val else ""

TOKEN = clean("TELEGRAM_TOKEN")
CHAT_ID = clean("TELEGRAM_CHAT_ID")

st.set_page_config(page_title="Khalil Audit", layout="centered")
st.title("🛡️ Khalil Trade: Отправка аудита")

# --- 2. ИНТЕРФЕЙС ---
with st.form("audit_form"):
    st.subheader("📋 Данные проверки")
    company = st.text_input("Название организации")
    expert = st.text_input("ФИО эксперта", value="Иван Рудой")
    
    col1, col2 = st.columns(2)
    with col1:
        pc_count = st.number_input("Кол-во ПК", min_value=0)
    with col2:
        servers = st.number_input("Серверы", min_value=0)
    
    st.divider()
    submit = st.form_submit_button("🚀 СФОРМИРОВАТЬ И ОТПРАВИТЬ")

# --- 3. ЛОГИКА ---
if submit:
    if not company or not expert:
        st.error("⚠️ Пожалуйста, заполните название компании и ФИО.")
    else:
        with st.status("Создание отчета и отправка...") as status:
            try:
                # Генерируем Excel в памяти
                output = BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Аудит"
                ws.append(["Параметр", "Значение"])
                ws.append(["Компания", company])
                ws.append(["Эксперт", expert])
                ws.append(["Кол-во ПК", pc_count])
                ws.append(["Серверы", servers])
                wb.save(output)
                
                # Отправка в Telegram
                url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
                files = {'document': (f"Audit_{company}.xlsx", output.getvalue())}
                payload = {
                    "chat_id": CHAT_ID,
                    "caption": f"✅ *НОВЫЙ АУДИТ*\n🏢 {company}\n👤 {expert}",
                    "parse_mode": "Markdown"
                }
                
                response = requests.post(url, data=payload, files=files, timeout=20)
                
                if response.ok:
                    status.update(label="✅ Успешно доставлено!", state="complete")
                    st.balloons()
                    st.success(f"Отчет по компании '{company}' отправлен в ваш Telegram.")
                else:
                    status.update(label="❌ Ошибка доставки", state="error")
                    st.error("Telegram не принял файл.")
                    st.json(response.json()) # Тут будет написано почему
                    
            except Exception as e:
                st.error(f"Сбой приложения: {e}")

st.caption(f"Бот: @Khaliltrade_bot | ID назначения: {CHAT_ID}")
