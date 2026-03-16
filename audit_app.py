import streamlit as st
import requests
from io import BytesIO
from openpyxl import Workbook

# --- ВСТАВЬТЕ ВАШИ ДАННЫЕ СЮДА НАПРЯМУЮ ДЛЯ ТЕСТА ---
TEST_TOKEN = "8757743843:AAELoPZlnUf0K5P0HYNxXyr0EswTfdqqD2o"
TEST_CHAT_ID = "ВАШ_ID_ИЗ_USERINFOBOT" # Например: "125320424"

st.set_page_config(page_title="Khalil Audit Final Test")
st.title("🛡️ Финальный тест (Прямое подключение)")

st.write("Этот тест игнорирует настройки Secrets и стучится в Telegram напрямую.")

company = st.text_input("Название компании", value="Тест Алматы")

if st.button("🚀 ОТПРАВИТЬ ПРЯМЫМ ЗАПРОСОМ"):
    if not company:
        st.error("Введите название")
    else:
        try:
            # Создаем файл
            output = BytesIO()
            wb = Workbook()
            ws = wb.active
            ws.append(["Тест связи", "Проверка"])
            ws.append(["Компания", company])
            wb.save(output)
            
            # Прямая отправка
            url = f"https://api.telegram.org/bot{TEST_TOKEN}/sendDocument"
            
            payload = {
                "chat_id": TEST_CHAT_ID,
                "caption": f"✅ Прямая проверка для {company}",
                "parse_mode": "Markdown"
            }
            files = {'document': ("Test_Report.xlsx", output.getvalue())}
            
            # Делаем запрос
            r = requests.post(url, data=payload, files=files, timeout=15)
            
            if r.ok:
                st.success("🎉 ПОЛУЧИЛОСЬ! Файл должен быть в Telegram.")
                st.balloons()
            else:
                st.error(f"Ошибка: {r.status_code}")
                st.json(r.json())
                st.info("Если снова 'chat not found', попробуйте переслать любое сообщение из вашего бота боту @userinfobot — точно ли ID совпадает?")
        except Exception as e:
            st.error(f"Ошибка кода: {e}")
