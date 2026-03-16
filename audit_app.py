import streamlit as st
import requests
import re
from io import BytesIO
from openpyxl import Workbook

# --- 1. ОЧИСТКА ---
def ultra_clean(key):
    val = st.secrets.get(key, "")
    if not val: return ""
    return re.sub(r"[^a-zA-Z0-9:]", "", str(val))

TOKEN = ultra_clean("TELEGRAM_TOKEN")
# Важно: CHAT_ID должен быть строкой для корректной отправки
CHAT_ID = ultra_clean("TELEGRAM_CHAT_ID")

st.set_page_config(page_title="Khalil Audit", page_icon="🛡️")

st.title("🛡️ Khalil Trade Audit")

# --- 2. ФОРМА ---
with st.form("main_form"):
    company = st.text_input("Название компании")
    name = st.text_input("Эксперт", value="Иван Рудой")
    pcs = st.number_input("Кол-во ПК", min_value=0)
    
    submit = st.form_submit_button("🚀 ОТПРАВИТЬ ОТЧЕТ")

# --- 3. ЛОГИКА ---
if submit:
    if not company:
        st.warning("Введите название компании")
    else:
        try:
            # Создаем Excel
            output = BytesIO()
            wb = Workbook()
            ws = wb.active
            ws.append(["Параметр", "Значение"])
            ws.append(["Компания", company])
            ws.append(["ФИО", name])
            ws.append(["ПК", pcs])
            wb.save(output)
            
            # Отправка
            url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
            
            # Мы принудительно превращаем CHAT_ID в строку и убираем лишнее
            payload = {
                "chat_id": str(CHAT_ID).strip(),
                "caption": f"📄 Новый аудит: {company}\n👤 Эксперт: {name}",
                "parse_mode": "Markdown"
            }
            files = {'document': (f"Audit_{company}.xlsx", output.getvalue())}
            
            r = requests.post(url, data=payload, files=files, timeout=15)
            
            if r.ok:
                st.success("✅ Отчет успешно доставлен в Telegram!")
                st.balloons()
            else:
                st.error("❌ Ошибка отправки")
                st.json(r.json())
                st.info(f"Бот пытался отправить на ID: `{CHAT_ID}`. Если это ваш ID, напишите боту @Khaliltrade_bot любое сообщение еще раз.")
        
        except Exception as e:
            st.error(f"Ошибка: {e}")

st.caption(f"v1.1 | Бот: @Khaliltrade_bot | ID: {CHAT_ID}")
