import streamlit as st
import pandas as pd
import os
import requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

# --- НАСТРОЙКИ TELEGRAM (из Secrets) ---
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ЛОГОТИП И КОНТАКТЫ ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### Мы поможем Вам стать лучше!")
st.divider()

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v6.2")

# --- ИНСТРУКЦИЯ ДЛЯ ПОЛЬЗОВАТЕЛЯ ---
with st.expander("📖 Руководство по проведению экспресс-аудита"):
    st.markdown("""
    ### Инструкция
    1. Заполните данные о компании.
    2. Ответьте на вопросы в каждом техническом блоке.
    3. Нажмите кнопку внизу для формирования экспертного отчета.
    4. Отчет будет автоматически направлен администратору и доступен вам для скачивания.
    """)

# --- 3. СБОР ДАННЫХ О КЛИЕНТЕ ---
st.sidebar.header("🏢 Данные заказчика")
client_info = {
    "Город": st.sidebar.text_input("Город"),
    "Наименование компании": st.sidebar.text_input("Наименование компании"),
    "ФИО контактного лица": st.sidebar.text_input("ФИО контактного лица"),
    "Сайт компании": st.sidebar.text_input("Сайт компании (URL)"),
    "Email": st.sidebar.text_input("Email"),
    "Контактный телефон": st.sidebar.text_input("Контактный телефон")
}

# --- 4. ОСНОВНЫЕ ВОПРОСЫ АУДИТА ---
data = []
score = 0
validation_errors = []

# Валидация Email (простая)
if client_info['Email'] and "@" not in client_info['Email']:
    validation_errors.append("Некорректный Email")

def ask_question(category, question, weight=10, justification=""):
    global score
    st.write(f"**{question}**")
    ans = st.radio(f"Выбор для: {question}", ["Да", "Нет", "В процессе внедрения"], key=question, label_visibility="collapsed")
    
    rec = ""
    status = "OK"
    is_risk = False
    
    if ans == "Да":
        score += weight
    elif ans == "Нет":
        status = "КРИТИЧЕСКИЙ РИСК"
        is_risk = True
        # Добавляем обоснование в рекомендацию
        rec = f"НЕОБХОДИМО ВНЕДРИТЬ. Обоснование: {justification}"
    else:
        score += (weight / 2)
        status = "ВНИМАНИЕ"
        rec = "Требуется ускорить завершение проекта и провести финальное тестирование."
    
    data.append({"Категория": category, "Вопрос": question, "Статус": status, "Рекомендация": rec, "is_risk": is_risk})

# --- БЛОК 1: ИНФРАСТРУКТУРА ---
with st.container():
    st.header("🌐 ИТ-Инфраструктура и Сервисы")
    col1, col2 = st.columns(2)
    with col1:
        ask_question("Почта", "Используется ли актуальная версия почтового сервера (напр. Exchange 2019+ или HCL Domino)?", 15, 
                     "Использование устаревших версий (напр. Exchange 2013/2016) несет риски эксплуатации неисправленных уязвимостей, на которые более не выпускаются патчи безопасности.")
    with col2:
        ask_question("Периметр", "Используется ли современный NGFW (напр. Check Point) вместо базовых роутеров?", 15, 
                     "Базовые роутеры не обеспечивают глубокую фильтрацию трафика (IPS/IDS) и защиту от современных угроз уровня приложений.")

# --- БЛОК 2: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ (ОБНОВЛЕННЫЙ) ---
st.divider()
with st.container():
    st.header("🛡️ Информационная безопасность")
    c1, c2 = st.columns(2)
    with c1:
        ask_question("Доступ", "Внедрена ли многофакторная аутентификация (MFA) для внешних доступов?", 20, 
                     "Пароли могут быть похищены через фишинг или перебор. MFA — единственный эффективный барьер, предотвращающий 99% атак на учетные данные.")
        ask_question("Привилегии", "Используется ли PAM-система для контроля действий администраторов?", 15, 
                     "Административные аккаунты — главная цель хакеров. Без PAM невозможно контролировать действия ИТ-персонала и подрядчиков внутри сети.")
    with c2:
        ask_question("Веб-защита", "Защищены ли веб-ресурсы и API с помощью WAAP/WAF решений?", 15, 
                     "Традиционные антивирусы и фаерволы не видят атаки на логику веб-приложения и попытки инъекций в базу данных.")
        ask_question("Антивирус", "Используется ли централизованный EDR/Antivirus на всех рабочих станциях?", 10, 
                     "Локальные антивирусы без центрального управления не позволяют вовремя обнаружить распространение вируса-шифровщика по сети.")

# --- 5. ГЕНЕРАЦИЯ EXCEL ОТЧЕТА ---
def make_expert_excel(c_info, audit_data, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Результаты аудита"
    
    # Стили
    header_font = Font(bold=True, color="FFFFFF", size=12)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Шапка данных клиента
    ws.append(["ОТЧЕТ ПО ТЕХНИЧЕСКОМУ АУДИТУ 2026"])
    ws.append([f"Компания: {c_info['Наименование компании']}"])
    ws.append([f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}"])
    ws.append([f"Общий уровень зрелости ИТ/ИБ: {final_score}%"])
    ws.append([])

    # Заголовки таблицы
    headers = ["Категория", "Объект аудита", "Текущий статус", "Экспертная рекомендация"]
    ws.append(headers)
    for cell in ws[6]:
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border

    # Данные
    curr_row = 7
    for item in audit_data:
        is_risk = item.pop('is_risk', False)
        ws.append(list(item.values()))
        
        # Подсветка рисков
        st_cell = ws.cell(row=curr_row, column=3)
        if is_risk: st_cell.font = Font(color="FF0000", bold=True)
        
        # Рамки для всех ячеек строки
        for col_idx in range(1, 5):
            ws.cell(row=curr_row, column=col_idx).border = border
            ws.cell(row=curr_row, column=col_idx).alignment = Alignment(wrap_text=True, vertical='top')
        curr_row += 1

    # Настройка ширины колонок
    for col, width in {'A': 20, 'B': 40, 'C': 20, 'D': 60}.items():
        ws.column_dimensions[col].width = width
    
    wb.save(output)
    return output.getvalue()

# --- 6. ФИНАЛЬНОЕ ДЕЙСТВИЕ ---
st.divider()

if validation_errors:
    st.error(f"🚨 Формирование отчета недоступно. Обнаружено ошибок: {len(validation_errors)}")

if st.button("📊 Сформировать экспертный отчет", disabled=len(validation_errors) > 0):
    mandatory = [client_info['Город'], client_info['Наименование компании'], client_info['ФИО контактного лица'], client_info['Email']]
    if not all(mandatory):
        st.error("⚠️ Пожалуйста, заполните все обязательные поля в боковой панели (Город, Компания, Контактное лицо, Email)!")
    else:
        with st.spinner("Анализируем данные и формируем отчет..."):
            f_score = min(int(score), 100)
            report_bytes = make_expert_excel(client_info, data, f_score)
            
            # Отправка в Telegram
            try:
                caption = (f"🚀 *Новый аудит Khalil Trade*\n"
                           f"🏢 *{client_info['Наименование компании']}*\n"
                           f"📍 {client_info['Город']}\n"
                           f"📊 Зрелость: {f_score}%\n"
                           f"👤 {client_info['ФИО контактного лица']}")
                
                requests.post(
                    f"https://api.telegram.org/bot{TOKEN}/sendDocument",
                    data={"chat_id": CHAT_ID, "caption": caption, "parse_mode": "Markdown"},
                    files={"document": ("Audit_Report.xlsx", report_bytes)}
                )
                st.success("✅ Отчет успешно сформирован и отправлен эксперту!")
                
                # Кнопка скачивания для пользователя
                st.download_button(
                    label="📥 Скачать ваш отчет (Excel)",
                    data=report_bytes,
                    file_name=f"Audit_{client_info['Наименование компании']}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Ошибка при отправке отчета: {e}")
