import streamlit as st
import requests
from io import BytesIO
from openpyxl import Workbook

# ВАШИ РАБОЧИЕ ДАННЫЕ (ПОДТВЕРЖДЕНО ТЕСТОМ)
TOKEN = "8589161723:AAG11MNF4-kpSP7Zonq74hPFKp9XhCWlQnQ"
MY_ID = "125320424"

st.set_page_config(page_title="Khalil Audit System", page_icon="🛡️")
st.title("🛡️ Система аудита Khalil Trade")

with st.form("audit_form"):
    st.subheader("Параметры проверки")
    company = st.text_input("Название организации")
    expert = st.text_input("Ведущий эксперт", value="Иван Рудой")
    pcs = st.number_input("Количество рабочих мест (АРМ)", min_value=0, step=1)
    
    st.divider()
    submit = st.form_submit_button("🚀 СФОРМИРОВАТЬ И ОТПРАВИТЬ ОТЧЕТ")

if submit:
    if not company:
        st.error("Пожалуйста, введите название организации!")
    else:
        with st.spinner("Генерация отчета и отправка в Telegram..."):
            try:
                # 1. Создание Excel-файла в памяти
                output = BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Audit Data"
                ws.append(["Параметр", "Значение"])
                ws.append(["Организация", company])
                ws.append(["Эксперт", expert])
                ws.append(["Кол-во АРМ", pcs])
                ws.append(["Дата аудита", "17.03.2026"])
                wb.save(output)
                
                # 2. Отправка через Telegram Bot API
                url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
                caption = (f"📊 *Новый аудит информационной безопасности*\n\n"
                           f"🏢 *Объект:* {company}\n"
                           f"👤 *Эксперт:* {expert}\n"
                           f"🖥️ *Кол-во АРМ:* {pcs}\n"
                           f"📅 *Дата:* 17.03.2026")
                
                files = {'document': (f"Audit_{company}_17_03.xlsx", output.getvalue())}
                payload = {"chat_id": MY_ID, "caption": caption, "parse_mode": "Markdown"}
                
                r = requests.post(url, data=payload, files=files)
                
                if r.ok:
                    st.success(f"Отчет для '{company}' успешно отправлен в ваш Telegram!")
                    st.balloons()
                else:
                    st.error(f"Ошибка при отправке: {r.text}")
            except Exception as e:
                st.error(f"Критическая ошибка: {e}")

st.caption("Developed by Ivan Rudoy | Almaty 2026")
