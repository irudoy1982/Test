import streamlit as st
import requests

# Ваш последний токен
TOKEN = "8757743843:AAELoPZlnUf0K5P0HYNxXyr0EswTfdqqD2o"

st.title("🔎 Поиск Ивана")

if st.button("КТО МНЕ ПИСАЛ?"):
    res = requests.get(f"https://api.telegram.org/bot{TOKEN}/getUpdates").json()
    if res.get("result"):
        # Показываем все сообщения, которые пришли боту
        for update in res["result"]:
            chat = update["message"]["chat"]
            st.write(f"✅ Нашел! Имя: {chat.get('first_name')}, Ваш ID: `{chat.get('id')}`")
    else:
        st.error("Бот всё еще никого не видит. Иван, напишите боту в Telegram любое слово ПРЯМО СЕЙЧАС.")
