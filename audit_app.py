import streamlit as st
import requests

TOKEN = st.secrets.get("TELEGRAM_TOKEN", "").strip().replace('"', '')
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID", "").strip().replace('"', '')

st.title("🛡️ Финальный тест связи")

if st.button("📨 ОТПРАВИТЬ ПРОСТО ТЕКСТ"):
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    payload = {
        "chat_id": CHAT_ID,
        "text": "🚀 Если вы это видите, то ID и Токен настроены ВЕРНО!"
    }
    r = requests.post(url, data=payload)
    if r.ok:
        st.success("ТЕКСТ ПРИШЕЛ! Значит, связь работает.")
    else:
        st.error(f"Ошибка: {r.text}")

if st.button("🔍 ПОСМОТРЕТЬ КТО ПИСАЛ (getUpdates)"):
    # Этот метод покажет всех, кто писал боту за последние 24 часа
    res = requests.get(f"https://api.telegram.org/bot{TOKEN}/getUpdates").json()
    st.json(res)
