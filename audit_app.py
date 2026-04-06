import streamlit as st
import pandas as pd
import os
import requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026 Expert", layout="wide", page_icon="🛡️")

# --- НАСТРОЙКИ TELEGRAM (из Secrets) ---
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ЛОГОТИП И ЗАГОЛОВОК ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade IT Audit & Consulting")

st.markdown("### Экспертная система оценки технологической зрелости")
st.divider()

# --- 3. СВЕДЕНИЯ О КЛИЕНТЕ И ОТРАСЛЬ ---
with st.sidebar:
    st.header("🏢 Сведения об организации")
    company_name = st.text_input("Наименование компании", placeholder="ООО 'Вектор'")
    
    # ПУНКТ 5: Отраслевая специфика
    industry = st.selectbox(
        "Сфера деятельности", 
        ["Финтех / Банки", "Ритейл / E-commerce", "Производство", "IT / Разработка", "Госсектор", "Другое"],
        help="Отрасль критически влияет на оценку рисков. Например, требования к ИБ в Финтехе значительно выше, чем в производстве."
    )
    
    contact_person = st.text_input("Контактное лицо")
    
    client_info = {
        "Наименование компании": company_name,
        "Сфера деятельности": industry,
        "Контактное лицо": contact_person,
        "Дата аудита": datetime.now().strftime("%d.%m.%Y")
    }

# --- 4. ОПРОСНИК С ДИНАМИЧЕСКИМИ ПОДСКАЗКАМИ (Пункт 4) ---
data = {}
score = 0
validation_errors = []

st.header("📋 Технический опросник")

col1, col2 = st.columns(2)

with col1:
    st.subheader("🌐 Инфраструктура и Сеть")
    
    data['1.1. Всего АРМ'] = st.number_input(
        "Общее количество рабочих станций (АРМ)", 
        min_value=0, value=10,
        help="Количество ПК и ноутбуков определяет масштаб сети. При >50 АРМ обязательна сегментация сети."
    )
    if data['1.1. Всего АРМ'] > 0: score += 5

    data['1.2.6. Коммутаторы L3'] = st.checkbox(
        "Используются L3-коммутаторы (Core/Aggregation)",
        help="L3-коммутаторы позволяют маршрутизировать трафик между VLAN. Это необходимо для безопасности и снижения широковещательного шторма."
    )
    if data['1.2.6. Коммутаторы L3']: score += 10

    data['Wi-Fi Точки доступа'] = st.number_input(
        "Количество точек доступа Wi-Fi", 
        min_value=0, value=0,
        help="Если точек более 5, экспертно рекомендуется использование аппаратного или программного контроллера."
    )
    
    data['Wi-Fi Контроллер'] = st.checkbox(
        "Наличие контроллера Wi-Fi",
        help="Контроллер обеспечивает бесшовный роуминг (переключение между точками без обрыва связи)."
    )
    if data['Wi-Fi Контроллер']: score += 5

with col2:
    st.subheader("🛡️ Информационная безопасность")
    
    data['MFA (Аутентификация)'] = st.checkbox(
        "Используется MFA/2FA (Второй фактор)",
        help="Многофакторная аутентификация для почты, VPN и админ-доступов. Это защита №1 от взлома."
    )
    if data['MFA (Аутентификация)']: score += 15

    data['1.2.7. NGFW'] = st.checkbox(
        "Наличие межсетевого экрана нового поколения (NGFW)",
        help="Решения вроде FortiGate, CheckPoint или UserGate. Проверяют не только порты, но и содержимое трафика."
    )
    if data['1.2.7. NGFW']: score += 15

    data['SIEM'] = st.checkbox(
        "Используется SIEM-система",
        help="Система сбора и анализа событий ИБ. Позволяет обнаружить атаку хакера, когда антивирусы молчат."
    )
    if data['SIEM']: score += 10

st.divider()

col3, col4 = st.columns(2)

with col3:
    st.subheader("💻 Серверы и Сервисы")
    
    data['1.3.2. Виртуальные серверы'] = st.number_input("Количество виртуальных серверов", min_value=0, value=0)
    data['1.3.1. Физические серверы'] = st.number_input("Количество физических серверов", min_value=0, value=0)
    
    data['Резервное копирование'] = st.checkbox(
        "Наличие системы резервного копирования (СРК)",
        help="Регулярное создание копий данных. Без СРК риск полной остановки бизнеса при аварии — 100%."
    )
    if data['Резервное копирование']: score += 15
    
    data['Хранение вне офиса'] = st.checkbox(
        "Копии хранятся в облаке или на другой площадке",
        help="Защита от пожара или кражи оборудования в основном офисе (правило 3-2-1)."
    )
    if data['Хранение вне офиса']: score += 5

with col4:
    st.subheader("🚀 Разработка и ИС")
    
    data['4.1. Разработчики'] = st.number_input(
        "Количество штатных разработчиков", 
        min_value=0, value=0,
        help="Если у вас есть разработка, требования к защите веб-приложений (WAF) и хранению кода (Git) становятся критичными."
    )
    
    data['WAF (Веб)'] = st.checkbox(
        "Используется WAF (Web Application Firewall)",
        help="Защищает ваши сайты и API от взлома (SQL-инъекции, XSS). Обязательно для E-commerce и Финтеха."
    )
    if data['WAF (Веб)']: score += 10

    data['4.3. Хранение кода (Git)'] = st.checkbox(
        "Используется централизованный Git (GitLab/GitHub)",
        help="Предотвращает потерю интеллектуальной собственности компании."
    )
    if data['4.3. Хранение кода (Git)']: score += 5

# --- 5. ФУНКЦИЯ ЭКСПЕРТНОЙ ГЕНЕРАЦИИ EXCEL (Логика v10.0) ---
def make_expert_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Report 2026"
    
    # Стили
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Заголовок
    ws.merge_cells('A1:E2')
    ws['A1'] = "СТРАТЕГИЧЕСКИЙ АУДИТ ИТ И ИБ (CONFORMITY NIST/ISO)"
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=14, color="1F4E78")

    # Сведения о клиенте
    curr_row = 4
    for k, v in c_info.items():
        ws.cell(row=curr_row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=curr_row, column=2, value=str(v))
        curr_row += 1
    
    ws.cell(row=curr_row, column=1, value="ИНДЕКС ЗРЕЛОСТИ:").font = Font(bold=True)
    score_cell = ws.cell(row=curr_row, column=2, value=f"{final_score}%")
    bg_col = "92D050" if final_score > 70 else "FFC000" if final_score > 40 else "FF7C80"
    score_cell.fill = PatternFill(start_color=bg_col, end_color=bg_col, fill_type="solid")
    
    curr_row += 3
    headers = ["Параметр", "Значение", "Статус", "Рекомендация (Best Practice)", "Стандарт (ISO/NIST)"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=curr_row, column=i, value=h)
        cell.fill = header_fill; cell.font = white_font
    curr_row += 1

    # Аналитические переменные
    industry = c_info.get("Сфера деятельности", "Другое")
    total_arm = results.get('1.1. Всего АРМ', 0)
    dev_count = results.get('4.1. Разработчики', 0)
    wifi_ap = results.get('Wi-Fi Точки доступа', 0)

    for k, v in results.items():
        status = "В норме"
        rec = "Риски минимальны."
        std = "N/A"
        val_str = str(v).lower()
        is_absent = "нет" in val_str or v is False or v == 0

        # --- ЭКСПЕРТНАЯ ЛОГИКА ---
        
        # SIEM и Отрасль
        if "SIEM" in k:
            std = "ISO 27001 (A.12.4)"
            if is_absent:
                if industry in ["Финтех / Банки", "IT / Разработка"]:
                    status = "КРИТИЧНО"; rec = f"Для отрасли '{industry}' отсутствие SIEM недопустимо. Риск потери лицензии."
                elif total_arm > 100:
                    status = "ВЫСОКИЙ РИСК"; rec = "Сеть более 100 АРМ требует автоматизированного мониторинга событий."
        
        # Wi-Fi и Роуминг
        elif "Точки доступа" in k and wifi_ap > 0:
            std = "IEEE 802.11ax"
            if wifi_ap > 5 and not results.get('Wi-Fi Контроллер'):
                status = "ВЫСОКИЙ РИСК"; rec = "Используется много точек без контроллера. Возможны обрывы связи при перемещении."
            elif wifi_ap > 0 and (total_arm / wifi_ap) > 25:
                status = "ВНИМАНИЕ"; rec = "Высокая плотность устройств на одну точку. Рекомендуется Wi-Fi 6."

        # WAF и Разработка
        elif "WAF" in k:
            std = "OWASP Top 10"
            if is_absent and (dev_count > 0 or industry == "Ритейл / E-commerce"):
                status = "КРИТИЧНО"; rec = "Публичные веб-ресурсы не защищены от взлома. Необходим WAF."

        # MFA
        elif "MFA" in k and is_absent:
            std = "NIST SP 800-63"; status = "КРИТИЧНО"; rec = "80% взломов происходят из-за кражи паролей. Срочно внедрить второй фактор."

        # Резервное копирование
        elif "Резервное копирование" in k:
            std = "ISO 27001 (A.12.3)"
            if is_absent:
                status = "КРИТИЧНО"; rec = "Отсутствие бэкапа — это риск полной гибели бизнеса при любой аварии."
        
        # Общий риск для пустых полей
        elif is_absent and status == "В норме":
            status = "РИСК"; rec = "Система отсутствует. Рекомендуется экспертная оценка целесообразности."

        # Запись строки
        row_vals = [k, str(v), status, rec, std]
        for col_idx, value in enumerate(row_vals, 1):
            cell = ws.cell(row=curr_row, column=col_idx, value=value)
            cell.border = border
            if col_idx == 3:
                if status == "КРИТИЧНО": cell.font = Font(color="FF0000", bold=True)
                elif status == "ВЫСОКИЙ РИСК": cell.font = Font(color="C00000", bold=True)
                elif status in ["ВНИМАНИЕ", "РИСК"]: cell.font = Font(color="FF8C00", bold=True)
            if col_idx == 4:
                cell.alignment = Alignment(wrapText=True, vertical='top')

        curr_row += 1

    # Ширина колонок
    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 50
    ws.column_dimensions['E'].width = 20

    wb.save(output)
    return output.getvalue(), datetime.now().strftime("%d.%m.%Y %H:%M")

# --- 6. ЗАВЕРШЕНИЕ И ВЫГРУЗКА ---
st.divider()

if not company_name:
    validation_errors.append("Введите наименование компании в боковом меню.")

if st.button("📊 Сформировать экспертный отчет", disabled=len(validation_errors) > 0):
    with st.spinner("Проводим глубокий анализ данных..."):
        f_score = min(score, 100)
        report_bytes, final_date = make_expert_excel(client_info, data, f_score)
        
        # Отправка в Telegram
        if TOKEN and CHAT_ID:
            try:
                caption = f"🚀 Новый аудит: {industry}\n🏢 {company_name}\n📊 Зрелость: {f_score}%"
                requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                            data={"chat_id": CHAT_ID, "caption": caption},
                            files={"document": ("Audit_Report.xlsx", report_bytes)})
            except:
                pass

        st.success(f"Анализ завершен! Индекс зрелости ИТ-инфраструктуры: {f_score}%")
        st.download_button(
            label="📥 Скачать экспертный отчет (Excel)",
            data=report_bytes,
            file_name=f"Audit_{company_name}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if validation_errors:
    for err in validation_errors:
        st.warning(err)
