import streamlit as st
import requests
from io import BytesIO
from openpyxl import Workbook

# Вставляем ваш НОВЫЙ токен напрямую
TOKEN = "8757743843:AAELoPZlnUf0K5P0HYNxXyr0EswTfdqqD2o"

st.set_page_config(page_title="Khalil Audit System", layout="centered")
st.title("🛡️ Система аудита Khalil Trade")

# --- СЕКЦИЯ ДИАГНОСТИКИ ---
with st.expander("🔍 ШАГ 1: ПОЛУЧИТЬ МОЙ ID (Если не работает)", expanded=True):
    st.write("1. Напишите любое сообщение боту @Khaliltrade_bot")
    if st.button("УЗНАТЬ МОЙ ID ИЗ ТЕЛЕГРАМ"):
        res = requests.get(f"https://api.telegram.org/bot{TOKEN}/getUpdates").json()
        if res.get("result"):
            # Берем ID из последнего сообщения
            chat_info = res["result"][-1]["message"]["chat"]
            real_id = chat_info["id"]
            st.success(f"✅ Бот вас увидел! Ваш ID: `{real_id}`")
            st.info("Скопируйте это число в поле ниже.")
        else:
            st.error("Бот пока не видит сообщений. Пожалуйста, напишите ему что-нибудь в Telegram прямо сейчас!")

# --- СЕКЦИЯ ОТПРАВКИ ---
st.divider()
target_id = st.text_input("Введите ваш ID для получения отчетов", value=st.secrets.get("TELEGRAM_CHAT_ID", ""))

with st.form("audit_form"):
    company = st.text_input("Название компании")
    expert = st.text_input("Эксперт", value="Иван Рудой")
    pcs = st.number_input("Количество ПК", min_value=0)
    
    submit = st.form_submit_button("🚀 СФОРМИРОВАТЬ И ОТПРАВИТЬ")

if submit:
    if not company or not target_id:
        st.error("Заполните название компании и ваш ID!")
    else:
        try:
            # Создаем файл
            output = BytesIO()
            wb = Workbook()
            ws = wb.active
            ws.append(["Параметр", "Значение"])
            ws.append(["Компания", company])
            ws.append(["Эксперт", expert])
            ws.append(["ПК", pcs])
            wb.save(output)
            
            # Отправка
            url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
            payload = {
                "chat_id": target_id.strip(),
                "caption": f"📊 *Отчет по аудиту*\n🏢 Компания: {company}\n👤 Эксперт: {expert}",
                "parse_mode": "Markdown"
            }
            files = {'document': (f"Audit_{company}.xlsx", output.getvalue())}
            
            r = requests.post(url, data=payload, files=files)
            
            if r.ok:
                st.success("🎉 Готово! Проверьте Telegram.")
                st.balloons()
            else:
                st.error("Ошибка при отправке:")
                st.json(r.json())
        except Exception as e:
            st.error(f"Ошибка: {e}")

st.caption("Ivan Rudoy | Almaty 2026")
