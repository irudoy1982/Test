import streamlit as st
import requests
import re
from io import BytesIO
from openpyxl import Workbook

# --- 1. ТОТАЛЬНАЯ ОЧИСТКА ---
def ultra_clean(key):
    # Берем значение из Secrets
    val = st.secrets.get(key, "")
    if not val:
        return ""
    # Оставляем только то, что может быть в токене и ID
    return re.sub(r"[^a-zA-Z0-9:]", "", str(val))

# Сразу получаем очищенные данные
TOKEN = ultra_clean("TELEGRAM_TOKEN")
CHAT_ID = ultra_clean("TELEGRAM_CHAT_ID")

st.set_page_config(page_title="Диагностика системы", layout="centered")

st.title("🛡️ Khalil Trade: Финальный тест")

# --- 2. ИНСТРУМЕНТ ПРОВЕРКИ ---
with st.expander("🛠 СКАНЕР ТОКЕНА (Нажмите здесь)", expanded=True):
    st.write("Ниже показано то, что видит код внутри вашего приложения:")
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Длина токена", len(TOKEN))
        st.write(f"Начало: `{TOKEN[:8]}...`")
    with col2:
        st.metric("Длина ID", len(CHAT_ID))
        st.write(f"ID: `{CHAT_ID}`")
    
    if st.button("🔌 ПРОВЕРИТЬ СВЯЗЬ С BOTFATHER"):
        try:
            # Самый простой запрос к Telegram: "Кто я?"
            test_url = f"https://api.telegram.org/bot{TOKEN}/getMe"
            r = requests.get(test_url, timeout=10)
            if r.ok:
                bot_info = r.json()
                st.success(f"✅ СВЯЗЬ ЕСТЬ! Имя бота: @{bot_info['result']['username']}")
            else:
                st.error(f"❌ СВЯЗИ НЕТ. Ошибка {r.status_code}")
                st.json(r.json())
        except Exception as e:
            st.error(f"Ошибка запроса: {e}")

st.divider()

# --- 3. ФОРМА ОТПРАВКИ ---
with st.form("debug_form"):
    company = st.text_input("Название компании (для теста)")
    name = st.text_input("Имя")
    submit = st.form_submit_button("🚀 ТЕСТОВАЯ ОТПРАВКА ФАЙЛА")

if submit:
    if not company or not name:
        st.warning("Заполните поля!")
    else:
        try:
            # Создаем микро-Excel
            output = BytesIO()
            wb = Workbook()
            ws = wb.active
            ws.append(["Тест", "Успешно"])
            ws.append(["Компания", company])
            wb.save(output)
            
            # Отправка
            url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
            files = {'document': (f"Test_{company}.xlsx", output.getvalue())}
            payload = {"chat_id": CHAT_ID, "caption": "Тест пройден!"}
            
            response = requests.post(url, data=payload, files=files, timeout=20)
            
            if response.ok:
                st.success("🎉 ФАЙЛ УШЕЛ! Проверьте Telegram.")
                st.balloons()
            else:
                st.error("Ошибка при отправке файла:")
                st.json(response.json())
        except Exception as e:
            st.error(f"Ошибка кода: {e}")

st.caption("Ivan Rudoy | 2026")
