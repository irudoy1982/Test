import streamlit as st
import requests
from io import BytesIO
from openpyxl import Workbook
from datetime import datetime

# --- ИНИЦИАЛИЗАЦИЯ И SECRETS ---
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

st.set_page_config(page_title="Khalil Audit Pro", page_icon="🛡️", layout="centered")

if not TOKEN or not CHAT_ID:
    st.error("⚠️ Ошибка: Настройте TELEGRAM_TOKEN и TELEGRAM_CHAT_ID в Secrets!")
    st.stop()

# --- ИНТЕРФЕЙС ---
st.title("🛡️ Khalil Audit Professional")
st.write(f"Сегодня: {datetime.now().strftime('%d.%m.%2026')}")

with st.form("main_audit_form"):
    # Секция 1: Общая информация
    st.subheader("📍 Информация об объекте")
    col1, col2 = st.columns(2)
    with col1:
        company = st.text_input("Название организации", placeholder="Напр. ТОО 'ТехноСервис'")
        expert = st.text_input("Ведущий эксперт", value="Иван Рудой")
    with col2:
        audit_type = st.selectbox("Тип проверки", ["Первичный аудит", "Плановый контроль", "Внеплановая проверка"])
        city = st.text_input("Город", value="Алматы")

    st.divider()

    # Секция 2: Техническая инвентаризация
    st.subheader("💻 ИТ-инфраструктура")
    c1, c2, c3 = st.columns(3)
    with c1:
        pcs = st.number_input("Кол-во АРМ (ПК)", min_value=0, step=1)
    with c2:
        servers = st.number_input("Серверы", min_value=0, step=1)
    with c3:
        printers = st.number_input("Периферия", min_value=0, step=1)

    st.divider()

    # Секция 3: Безопасность и ПО
    st.subheader("🔒 Кибербезопасность")
    os_versions = st.multiselect("Операционные системы", ["Windows 10/11", "Windows Server", "Linux", "macOS"])
    antivirus = st.radio("Антивирусная защита:", ["Установлена (Актуальна)", "Требует обновления", "Отсутствует"], horizontal=True)
    backup_status = st.checkbox("Резервное копирование настроено?")
    
    st.divider()

    # Секция 4: Заметки
    st.subheader("📝 Заключение")
    notes = st.text_area("Замечания и рекомендации эксперта", placeholder="Опишите выявленные уязвимости или требования...")

    submit = st.form_submit_button("🚀 СФОРМИРОВАТЬ И ОТПРАВИТЬ ОТЧЕТ")

# --- ЛОГИКА ОБРАБОТКИ ---
if submit:
    if not company:
        st.warning("Пожалуйста, введите название организации.")
    else:
        with st.status("Генерация отчета...") as status:
            try:
                # 1. Создание Excel-файла
                output = BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Результаты аудита"

                # Заполнение данных (заголовки и значения)
                rows = [
                    ["ПАРАМЕТР", "ЗНАЧЕНИЕ"],
                    ["Организация", company],
                    ["Дата аудита", datetime.now().strftime('%d.%m.%Y')],
                    ["Тип проверки", audit_type],
                    ["Город", city],
                    ["Эксперт", expert],
                    ["---", "---"],
                    ["Количество АРМ", pcs],
                    ["Количество серверов", servers],
                    ["Периферия", printers],
                    ["---", "---"],
                    ["ОС в сети", ", ".join(os_versions)],
                    ["Антивирус", antivirus],
                    ["Бэкап настроен", "Да" if backup_status else "Нет"],
                    ["---", "---"],
                    ["Замечания", notes]
                ]

                for r in rows:
                    ws.append(r)
                
                # Косметическое расширение колонок
                ws.column_dimensions['A'].width = 25
                ws.column_dimensions['B'].width = 45
                wb.save(output)

                # 2. Отправка в Telegram
                url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
                caption = (f"📑 *ОТЧЕТ ПО АУДИТУ*\n\n"
                           f"🏢 *Объект:* {company}\n"
                           f"👤 *Эксперт:* {expert}\n"
                           f"🖥️ *Инфраструктура:* {pcs} ПК, {servers} Серв.\n"
                           f"🛡️ *Антивирус:* {antivirus}\n"
                           f"📅 *Дата:* {datetime.now().strftime('%d.%m.%Y')}")

                files = {'document': (f"Audit_Report_{company}.xlsx", output.getvalue())}
                payload = {"chat_id": CHAT_ID, "caption": caption, "parse_mode": "Markdown"}

                response = requests.post(url, data=payload, files=files)

                if response.ok:
                    status.update(label="✅ Отчет доставлен в Telegram!", state="complete")
                    st.balloons()
                    st.success(f"Файл для {company} успешно отправлен Ивану.")
                else:
                    st.error(f"Ошибка API: {response.text}")

            except Exception as e:
                st.error(f"Техническая ошибка: {e}")

st.caption("Ivan Rudoy | IT Security Almaty 2026")
