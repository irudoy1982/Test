import streamlit as st
import requests
import os
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ 2026", layout="wide", page_icon="🛡️")

# Извлекаем секреты
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ЛОГОТИП И ШАПКА ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=250)
else:
    st.title("Khalil Trade | IT Audit")

st.markdown("### Экспертный анализ инфраструктуры")
st.divider()

# --- 3. ИНФОРМАЦИЯ О КЛИЕНТЕ ---
st.header("🏢 Информация о компании")
c1, c2 = st.columns(2)
with c1:
    company_name = st.text_input("Наименование компании", key="comp_name")
    contact_person = st.text_input("Контактное лицо (ФИО)", key="cont_pers")
with c2:
    email = st.text_input("Контактный email", key="cont_email")
    phone = st.text_input("Контактный телефон", key="cont_phone")

client_info = {
    "Компания": company_name,
    "Лицо": contact_person,
    "Телефон": phone,
    "Email": email
}

st.divider()

# --- 4. ТЕХНИЧЕСКИЙ АУДИТ ---
st.header("📋 Технический аудит")
data = {}
score = 0

with st.expander("Инфраструктура и Безопасность", expanded=True):
    data['АРМ'] = st.number_input("Кол-во АРМ (шт):", min_value=0, step=1)
    data['Серверы'] = st.number_input("Кол-во серверов:", min_value=0, step=1)
    
    ib_tasks = {
        "Резервное копирование": 25,
        "DLP (Защита данных)": 15,
        "EDR/Antimalware": 20,
        "NGFW (Межсетевой экран)": 15,
        "PAM (Контроль доступа)": 15,
        "WAF (Защита сайтов)": 10
    }
    
    for task, pts in ib_tasks.items():
        if st.checkbox(task, key=f"chk_{task}"):
            v_name = st.text_input(f"Вендор {task}:", key=f"v_{task}", placeholder="Укажите решение")
            data[task] = f"Да ({v_name if v_name else 'не указан'})"
            score += pts
        else:
            data[task] = "Нет"

# --- 5. ФУНКЦИЯ ГЕНЕРАЦИИ EXCEL ---
def make_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Report"
    
    # Оформление
    h_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    f_white = Font(color="FFFFFF", bold=True)
    side = Side(style='thin')
    brd = Border(left=side, right=side, top=side, bottom=side)

    ws.merge_cells('A1:D1')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ 2026"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center')

    ws['A3'] = "КОМПАНИЯ:"; ws['B3'] = c_info['Компания']
    ws['A4'] = "ИНДЕКС ЗРЕЛОСТИ:"; ws['B4'] = f"{final_score}/100"
    
    headers = ["Параметр", "Значение", "Статус", "Рекомендация"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=6, column=i, value=h)
        cell.fill = h_fill; cell.font = f_white; cell.border = brd

    row_num = 7
    for k, v in results.items():
        ws.cell(row=row_num, column=1, value=k).border = brd
        ws.cell(row=row_num, column=2, value=str(v)).border = brd
        status = "РИСК" if "Нет" in str(v) or v == 0 else "ОК"
        ws.cell(row=row_num, column=3, value=status).border = brd
        ws.cell(row=row_num, column=4, value="Требуется проверка" if status == "РИСК" else "В норме").border = brd
        row_num += 1

    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['D'].width = 40
    wb.save(output)
    return output.getvalue()

# --- 6. ФУНКЦИЯ ОТПРАВКИ В TELEGRAM ---
def send_telegram_diagnostics(c_info, final_score, excel_bytes):
    url = f"https://api.telegram.org/bot{TOKEN}"
    
    # ПРОВЕРКА 1: Текстовое сообщение
    st.write("⏳ Отправка текстового уведомления...")
    msg = (f"🔔 *НОВЫЙ АУДИТ*\n\n"
           f"🏢 *Компания:* {c_info['Компания']}\n"
           f"👤 *Контакт:* {c_info['Лицо']}\n"
           f"📊 *Результат:* {final_score}/100")
    
    try:
        r_text = requests.post(f"{url}/sendMessage", 
                               data={"chat_id": CHAT_ID, "text": msg, "parse_mode": "Markdown"},
                               timeout=15)
        if not r_text.ok:
            return f"Ошибка Telegram (Текст): {r_text.text}"
        st.write("✅ Текст доставлен!")

        # ПРОВЕРКА 2: Файл Excel
        st.write("⏳ Отправка файла Excel...")
        files = {'document': (f"Audit_{c_info['Компания']}.xlsx", excel_bytes)}
        r_file = requests.post(f"{url}/sendDocument", 
                               data={"chat_id": CHAT_ID}, 
                               files=files,
                               timeout=30)
        
        if not r_file.ok:
            return f"Ошибка Telegram (Файл): {r_file.text}"
        st.write("✅ Файл доставлен!")
        
        return True

    except requests.exceptions.Timeout:
        return "Ошибка: Превышено время ожидания (Timeout). Попробуйте еще раз."
    except Exception as e:
        return f"Критический сбой: {str(e)}"

# --- 7. КНОПКА ЗАПУСКА ---
st.divider()
if st.button("🚀 Сформировать и отправить отчет"):
    if not company_name:
        st.error("Укажите название компании!")
    elif not TOKEN or not CHAT_ID:
        st.error("Настройки Telegram (Secrets) не заполнены!")
    else:
        f_score = min(score, 100)
        excel_data = make_excel(client_info, data, f_score)
        
        # Индикация процесса
        with st.status("Выполнение процесса...", expanded=True) as status:
            result = send_telegram_diagnostics(client_info, f_score, excel_data)
            
            if result is True:
                status.update(label="✅ Все данные успешно отправлены!", state="complete", expanded=False)
                st.success("Отчет в вашем Telegram!")
                st.balloons()
                st.download_button("📥 Скачать копию Excel", excel_data, f"Audit_{company_name}.xlsx")
            else:
                status.update(label="❌ Ошибка отправки", state="error", expanded=True)
                st.error(result)
                st.info("💡 Проверьте, что вы нажали 'Start' в боте и Chat ID указан верно.")

st.caption("Developed by Ivan Rudoy | Khalil Trade 2026")
