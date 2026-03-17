import streamlit as st
import requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# --- НАСТРОЙКИ И БЕЗОПАСНОСТЬ ---
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
MY_ID = st.secrets.get("TELEGRAM_CHAT_ID")

st.set_page_config(page_title="Khalil Audit Pro", page_icon="🛡️", layout="centered")

# Проверка конфигурации
if not TOKEN or not MY_ID:
    st.error("⚠️ Настройте Secrets (TELEGRAM_TOKEN и TELEGRAM_CHAT_ID) в панели Streamlit!")
    st.stop()

# --- ИНТЕРФЕЙС ---
st.title("🛡️ Khalil Audit Professional")
st.info("Заполните данные аудита. После отправки вы получите готовый Excel-отчет в Telegram.")

with st.form("full_audit_form"):
    # Основная информация
    st.subheader("📍 Общие сведения")
    col1, col2 = st.columns(2)
    with col1:
        company = st.text_input("Название организации", placeholder="ТОО 'Пример'")
        expert = st.text_input("Эксперт", value="Иван Рудой")
    with col2:
        audit_date = st.date_input("Дата проведения")
        location = st.text_input("Город/Объект", value="Алматы")

    st.divider()

    # Технические показатели
    st.subheader("💻 Техническая инвентаризация")
    c1, c2, c3 = st.columns(3)
    with c1:
        pcs = st.number_input("Кол-во ПК (АРМ)", min_value=0, step=1)
    with c2:
        servers = st.number_input("Серверы", min_value=0, step=1)
    with c3:
        network_nodes = st.number_input("Сетевое обор.", min_value=0, step=1)

    st.divider()

    # Кибербезопасность
    st.subheader("🔒 Проверка систем защиты")
    pam_status = st.selectbox("Система PAM (Управление доступом)", ["Не внедрено", "В процессе", "Активно", "Требует обновления"])
    dlp_status = st.selectbox("Система DLP (Защита данных)", ["Отсутствует", "Внедрена частично", "Активна"])
    antivirus = st.radio("Антивирусная защита актуальна?", ["Да", "Нет", "Частично"], horizontal=True)
    
    notes = st.text_area("Дополнительные замечания и рекомендации")

    # Кнопка отправки
    st.markdown("---")
    submit = st.form_submit_button("🚀 СФОРМИРОВАТЬ И ОТПРАВИТЬ ПОЛНЫЙ ОТЧЕТ")

# --- ЛОГИКА ОБРАБОТКИ ---
if submit:
    if not company:
        st.warning("Пожалуйста, укажите название организации.")
    else:
        with st.spinner("Создание профессионального отчета..."):
            try:
                # 1. Генерация Excel
                output = BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Результаты аудита"

                # Стилизация заголовков
                bold_font = Font(bold=True)
                center_align = Alignment(horizontal='center')

                # Заполнение данных
                data = [
                    ["Параметр", "Значение"],
                    ["Организация", company],
                    ["Дата", str(audit_date)],
                    ["Город", location],
                    ["Эксперт", expert],
                    ["---", "---"],
                    ["Кол-во ПК (АРМ)", pcs],
                    ["Серверы", servers],
                    ["Сетевые узлы", network_nodes],
                    ["---", "---"],
                    ["Статус PAM", pam_status],
                    ["Статус DLP", dlp_status],
                    ["Антивирус", antivirus],
                    ["---", "---"],
                    ["Замечания", notes]
                ]

                for row in data:
                    ws.append(row)

                # Минимальное форматирование
                for cell in ws["A"]: cell.font = bold_font
                ws.column_dimensions['A'].width = 25
                ws.column_dimensions['B'].width = 40
                
                wb.save(output)

                # 2. Отправка в Telegram
                url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
                caption = (f"📑 *ПОЛНЫЙ ОТЧЕТ ПО АУДИТУ*\n\n"
                           f"🏢 *Объект:* {company}\n"
                           f"👤 *Эксперт:* {expert}\n"
                           f"📅 *Дата:* {audit_date}\n"
                           f"📍 *Локация:* {location}\n\n"
                           f"💻 *Инфраструктура:* {pcs} ПК, {servers} Серв.\n"
                           f"🔐 *PAM:* {pam_status}")

                files = {'document': (f"Audit_Report_{company}.xlsx", output.getvalue())}
                payload = {"chat_id": MY_ID, "caption": caption, "parse_mode": "Markdown"}
                
                r = requests.post(url, data=payload, files=files)

                if r.ok:
                    st.success(f"Отчет для {company} отправлен успешно!")
                    st.balloons()
                else:
                    st.error(f"Ошибка Telegram: {r.text}")

            except Exception as e:
                st.error(f"Произошла ошибка: {e}")

st.caption(f"Khalil Audit Tool | {location} 2026")
