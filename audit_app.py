import streamlit as st
import requests
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="ИТ-Аудит 2026", layout="centered", page_icon="🛡️")

# Функция тотальной очистки (убирает пробелы, кавычки и невидимые знаки)
def force_clean(text):
    if text:
        return re.sub(r"[^a-zA-Z0-9:]", "", str(text))
    return ""

# Получаем данные и СРАЗУ чистим
TOKEN = force_clean(st.secrets.get("TELEGRAM_TOKEN", ""))
CHAT_ID = force_clean(st.secrets.get("TELEGRAM_CHAT_ID", ""))

# --- 2. ИНТЕРФЕЙС ---
st.title("🛡️ Khalil Trade | IT Audit")
st.write("Система готова к отправке отчета.")

# Используем st.form, чтобы данные не терялись при вводе
with st.form("final_form"):
    st.subheader("📋 Данные для аудита")
    
    comp_name = st.text_input("Название компании", placeholder="Напр: ТОО 'Алма-Ата'")
    rep_name = st.text_input("Ваше имя")
    
    st.divider()
    
    st.subheader("⚙️ Технические параметры")
    col1, col2 = st.columns(2)
    with col1:
        pc = st.number_input("Кол-во ПК", min_value=0, step=1)
        srv = st.number_input("Кол-во серверов", min_value=0, step=1)
    with col2:
        is_backup = st.checkbox("Резервное копирование")
        is_av = st.checkbox("Антивирусная защита")

    submit = st.form_submit_button("🚀 ОТПРАВИТЬ В TELEGRAM")

# --- 3. ЛОГИКА ---
if submit:
    if not comp_name or not rep_name:
        st.warning("⚠️ Пожалуйста, заполните название компании и имя!")
    else:
        with st.status("Связь с Telegram...") as status:
            try:
                # Генерируем Excel
                output = BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Audit"
                
                # Заголовки
                ws.append(["Параметр", "Значение"])
                ws.append(["Компания", comp_name])
                ws.append(["Представитель", rep_name])
                ws.append(["ПК", pc])
                ws.append(["Серверы", srv])
                ws.append(["Бэкап", "Есть" if is_backup else "Нет"])
                ws.append(["Антивирус", "Есть" if is_av else "Нет"])
                
                # Немного оформления
                ws['A1'].font = Font(bold=True)
                ws['B1'].font = Font(bold=True)
                
                wb.save(output)
                file_data = output.getvalue()

                # Отправка
                url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
                payload = {
                    "chat_id": CHAT_ID,
                    "caption": f"🔔 *НОВЫЙ АУДИТ*\n🏢 Организация: {comp_name}\n👤 Эксперт: {rep_name}\n💻 Инфраструктура: {pc} ПК, {srv} Серв.",
                    "parse_mode": "Markdown"
                }
                files = {'document': (f"Audit_{comp_name}.xlsx", file_data)}
                
                r = requests.post(url, data=payload, files=files, timeout=20)
                
                if r.ok:
                    status.update(label="✅ Отчет успешно отправлен!", state="complete")
                    st.balloons()
                    st.success("Проверьте Telegram!")
                else:
                    status.update(label="❌ Ошибка Telegram", state="error")
                    st.error(f"Ответ сервера: {r.text}")
                    st.info(f"Проверьте, что в Secrets ID указан без кавычек. Ваш текущий ID: {CHAT_ID}")
            
            except Exception as e:
                st.error(f"Ошибка приложения: {e}")

st.caption(f"App Version 1.0.4 | User: Ivan Rudoy")
