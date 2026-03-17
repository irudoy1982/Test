import streamlit as st
import requests
from io import BytesIO
from openpyxl import Workbook

# Безопасное получение данных из Secrets
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
MY_ID = st.secrets.get("TELEGRAM_CHAT_ID")

st.set_page_config(page_title="Khalil Audit", page_icon="🛡️")
st.title("🛡️ Khalil Audit System")

# Проверка наличия ключей
if not TOKEN or not MY_ID:
    st.error("⚠️ Ошибка: Настройте TELEGRAM_TOKEN и TELEGRAM_CHAT_ID в Secrets!")
    st.stop()

with st.form("audit_form"):
    company = st.text_input("Название организации")
    expert = st.text_input("Эксперт", value="Иван Рудой")
    pcs = st.number_input("Количество АРМ", min_value=0, step=1)
    submit = st.form_submit_button("🚀 ОТПРАВИТЬ ОТЧЕТ")

if submit:
    if not company:
        st.warning("Введите название компании")
    else:
        try:
            output = BytesIO()
            wb = Workbook()
            ws = wb.active
            ws.append(["Параметр", "Значение"])
            ws.append(["Организация", company])
            ws.append(["Эксперт", expert])
            ws.append(["Кол-во АРМ", pcs])
            ws.append(["Дата", "17.03.2026"])
            wb.save(output)

            url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
            caption = f"📊 *Отчет сформирован*\n🏢 Объект: {company}\n👤 Эксперт: {expert}"
            
            files = {'document': (f"Audit_{company}.xlsx", output.getvalue())}
            payload = {"chat_id": MY_ID, "caption": caption, "parse_mode": "Markdown"}
            
            r = requests.post(url, data=payload, files=files)
            
            if r.ok:
                st.success("✅ Файл успешно отправлен через Secrets!")
                st.balloons()
            else:
                st.error(f"Ошибка: {r.text}")
        except Exception as e:
            st.error(f"Сбой: {e}")

st.caption("Almaty 2026 | Secured Version")
