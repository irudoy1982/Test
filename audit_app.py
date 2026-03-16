import streamlit as st
import requests

# Берем токен из ваших Secrets
TOKEN = st.secrets.get("TELEGRAM_TOKEN", "").strip().replace('"', '')

st.title("🔎 Кто я на самом деле?")

if st.button("УЗНАТЬ ИМЯ БОТА"):
    if not TOKEN:
        st.error("Токен не найден в Secrets!")
    else:
        try:
            # Запрос к Telegram, чтобы узнать инфо о боте
            res = requests.get(f"https://api.telegram.org/bot{TOKEN}/getMe").json()
            if res.get("ok"):
                bot_info = res["result"]
                st.success("✅ Токен рабочий!")
                st.write(f"**Имя бота:** {bot_info.get('first_name')}")
                st.write(f"**Username:** @{bot_info.get('username')}")
                st.info("Внимание: Вы должны писать сообщения ИМЕННО ЭТОМУ боту, который указан выше!")
            else:
                st.error("❌ Telegram не принимает этот токен (401 Unauthorized)")
        except Exception as e:
            st.error(f"Ошибка: {e}")
