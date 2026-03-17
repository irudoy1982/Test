import streamlit as st
import pandas as pd
import os
import requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
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

st.markdown("### Мы поможем Вам стать лучше!**")
st.divider()

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026)")

data = {}
client_info = {}
score = 0

# --- ШАПКА: ИНФОРМАЦИЯ О КЛИЕНТЕ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город:*")
    filling_date = st.date_input("Дата заполнения:*", value=datetime.now())
    client_info['Дата заполнения'] = filling_date.strftime("%d.%m.%Y")
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    client_info['Сайт компании'] = st.text_input("Сайт компании (необязательно):")

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    client_info['Контактный телефон'] = st.text_input("Контактный телефон:*")

st.divider()

# --- ТЕХНИЧЕСКИЕ БЛОКИ (БЕЗ ИЗМЕНЕНИЙ В ЛОГИКЕ) ---
# Блок 1: Инфраструктура
st.header("Блок 1: Техническая часть")
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1, key="total_arm_val")
data['1.1. Всего АРМ'] = total_arm
selected_os_arm = st.multiselect("Выберите ОС на АРМ:", ["Windows", "Linux", "macOS", "Другое"], key="ms_arm_list")
if selected_os_arm:
    for os_item in selected_os_arm:
        count_arm = st.number_input(f"Количество АРМ на {os_item}:", min_value=0, step=1, key=f"arm_cnt_{os_item}")
        data[f"ОС АРМ ({os_item})"] = count_arm

st.write("---")
st.subheader("1.2. Серверы")
col_s1, col_s2 = st.columns(2)
with col_s1:
    phys_servers = st.number_input("Количество физических серверов:", min_value=0, step=1, key="phys_srv")
    data['1.2. Физические серверы'] = phys_servers
with col_s2:
    virt_servers = st.number_input("Количество виртуальных серверов:", min_value=0, step=1, key="virt_srv")
    data['1.2. Виртуальные серверы'] = virt_servers

selected_os_srv = st.multiselect("Выберите ОС серверов:", ["Windows Server", "Linux", "Unix", "Другое"], key="ms_srv_list")
if selected_os_srv:
    for os_s in selected_os_srv:
        count_srv = st.number_input(f"Количество серверов на {os_s}:", min_value=0, step=1, key=f"srv_cnt_{os_s}")
        data[f"ОС Сервера ({os_s})"] = count_srv

st.write("---")
col_v1, col_v2 = st.columns(2)
with col_v1:
    st.subheader("1.3. Виртуализация")
    data['1.3. Виртуализация'] = st.multiselect("Системы виртуализации:", ["VMware", "Hyper-V", "Proxmox", "KVM", "Другое", "Нет"], key="virt_sys")
with col_v2:
    st.subheader("1.4. Почтовая система")
    data['1.4. Почта'] = st.selectbox("Тип почты:", ["Exchange (On-Prem)", "Microsoft 365", "Google Workspace", "Yandex/Mail.ru Cloud", "Собственный сервер", "Нет"], key="mail_sys")

st.write("---")
st.subheader("1.5. Внутренние Информационные системы")
has_is = st.checkbox("Есть ли внутренние Информационные системы (1C, ERP, CRM)?", key="is_15_chk")
data['1.5. Внутренние ИС'] = st.text_input("Перечислите их:", key="is_input_field") if has_is else "Нет"

st.subheader("1.6. Система мониторинга")
has_mon = st.checkbox("Есть ли система мониторинга?", key="mon_16_chk")
data['1.6. Мониторинг'] = st.selectbox("Выберите:", ["Zabbix", "Nagios", "PRTG", "Prometheus", "Другое"], key="mon_select") if has_mon else "Нет"

# Блок 2: Сеть
st.header("Блок 2: Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_block_toggle"):
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink"]
    c_n1, c_n2 = st.columns(2)
    with c_n1:
        data['2.1. Основной канал'] = st.selectbox("Тип (Основной):", net_types, key="main_net_type")
        data['2.1. Скорость (mbit/s)'] = st.number_input("Скорость:", min_value=0, key="main_net_speed")
    with c_n2:
        data['2.1. Резервный канал'] = st.selectbox("Тип (Резервный):", ["Нет"] + net_types, key="res_net_type")
    
    data['2.4. NGFW'] = st.text_input("Вендор NGFW:", key="ngfw_vendor")
    if data['2.4. NGFW']: score += 20

    if st.checkbox("Используется Wi-Fi?", key="wifi_usage_chk"):
        data['2.5. Контроллер'] = st.text_input("Модель контроллера:", key="wifi_ctrl")
        data['2.5. Число точек'] = st.number_input("Кол-во точек:", min_value=0, key="wifi_ap_cnt")

# Блок 3: ИБ
st.header("Блок 3: Информационная Безопасность")
if st.toggle("Есть отдел ИБ", key="ib_block_toggle"):
    ib_list = {
        "DLP (Защита от утечек)": 15, "PAM (Контроль доступа)": 10, "SIEM/SOC": 20, 
        "WAF (Защита Web)": 10, "EDR/Antimalware": 15, "Резервное копирование": 20
    }
    for label, pts in ib_list.items():
        c_ib1, c_ib2 = st.columns([1, 2])
        if c_ib1.checkbox(label, key=f"chk_{label}"):
            v_n = c_ib2.text_input(f"Вендор {label}:", key=f"v_{label}")
            data[label] = f"Да ({v_n if v_n else 'не указан'})"
            score += pts
        else:
            data[label] = "Нет"
    
    if st.checkbox("3.7. Другие системы защиты", key="ib_other_toggle"):
        data['3.7. Прочее ИБ'] = st.text_area("Описание:", key="ib_other_input")

# Блок 4: Web
st.header("Блок 4: Web-ресурсы")
if st.toggle("Есть свои Web-ресурсы", key="web_toggle"):
    data['4.1. Хостинг'] = st.selectbox("Хостинг:", ["Собственный ЦОД", "Облако (KZ)", "Облако (Global)", "Нет"], key="web_host")
    data['4.2. CMS'] = st.text_input("CMS:", key="web_cms")
    data['4.4. Frontend'] = st.multiselect("Frontend серверы:", ["Nginx", "Apache", "IIS", "LiteSpeed", "Caddy", "Cloudflare"], key="web_front")

# Блок 5: Разработка
st.header("Блок 5: Разработка")
if st.toggle("Своя разработка", key="dev_toggle"):
    data['5.1. Разработчики'] = st.number_input("Кол-во разработчиков:", min_value=0, key="dev_cnt")
    data['5.3. CI/CD'] = st.checkbox("Используется CI/CD?", key="dev_cicd")
    data['5.4. Контейнеры'] = st.text_input("Технологии (Docker/K8s):", key="dev_cont")


# --- ЭКСЕЛЬ ГЕНЕРАЦИЯ (НОВАЯ ВЕРСИЯ С РЕКОМЕНДАЦИЯМИ) ---
def make_expert_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Khalil Audit Report"

    # Стили
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    bold_font = Font(bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # 1. Заголовок и Лого
    ws.merge_cells('A1:D2')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ ПО ИТ И ИБ (2026)"
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=16, color="1F4E78")

    if os.path.exists("logo.png"):
        try:
            img = OpenpyxlImage("logo.png")
            img.height = 60; img.width = 180
            ws.add_image(img, 'D1')
        except: pass

    # 2. Инфо о клиенте
    current_row = 4
    for k, v in c_info.items():
        ws.cell(row=current_row, column=1, value=k).font = bold_font
        ws.cell(row=current_row, column=2, value=str(v))
        current_row += 1

    # 3. Скоринг
    current_row += 1
    ws.cell(row=current_row, column=1, value="ИНДЕКС ТЕХНИЧЕСКОЙ ЗРЕЛОСТИ:").font = bold_font
    score_cell = ws.cell(row=current_row, column=2, value=f"{final_score}%")
    bg_color = "92D050" if final_score > 70 else "FFC000" if final_score > 40 else "FF7C80"
    score_cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    score_cell.font = bold_font
    score_cell.alignment = Alignment(horizontal='center')

    # 4. Таблица данных
    current_row += 2
    headers = ["Параметр", "Значение", "Статус", "Рекомендация эксперта Khalil Trade"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=current_row, column=i, value=h)
        cell.fill = header_fill; cell.font = white_font
        cell.alignment = Alignment(horizontal='center')

    # Логика рекомендаций
    rec_map = {
        "Нет": "Требуется внедрение для минимизации рисков.",
        "Резервное копирование": "Критично! Настроить схему 3-2-1 с хранением вне офиса.",
        "NGFW": "Рекомендуется NGFW для глубокого анализа трафика.",
        "DLP": "Необходимо для предотвращения утечек конфиденциальных данных.",
        "SIEM": "Рекомендуется для централизованного сбора логов и выявления атак.",
        "CI/CD": "Внедрение автоматизации ускорит выпуск и повысит качество кода."
    }

    current_row += 1
    for k, v in results.items():
        ws.cell(row=current_row, column=1, value=k).border = border
        ws.cell(row=current_row, column=2, value=str(v)).border = border
        
        # Определяем статус и рекомендацию
        val_str = str(v)
        status = "В норме"
        recommendation = "Поддерживать текущее состояние."

        if "Нет" in val_str or val_str == "0" or val_str == "[]":
            status = "РИСК"
            recommendation = rec_map.get(k, "Рассмотреть возможность внедрения.")
            st_cell = ws.cell(row=current_row, column=3, value=status)
            st_cell.font = Font(color="FF0000", bold=True)
        else:
            ws.cell(row=current_row, column=3, value=status)
            # Специальные советы для того, что УЖЕ есть
            if "Резервное копирование" in k: recommendation = "Регулярно проводить тестовое восстановление."
            if "NGFW" in k: recommendation = "Проверить актуальность подписок на сигнатуры."

        ws.cell(row=current_row, column=4, value=recommendation).border = border
        ws.cell(row=current_row, column=3).border = border
        current_row += 1

    # Ширина колонок
    dims = {'A': 35, 'B': 30, 'C': 15, 'D': 60}
    for col, width in dims.items(): ws.column_dimensions[col].width = width
    
    wb.save(output)
    return output.getvalue()


# --- ФИНАЛ И ОТПРАВКА ---
st.divider()
if st.button("📊 Сформировать экспертный отчет", key="final_btn_report"):
    mandatory = [client_info['Город'], client_info['Наименование компании'], client_info['ФИО контактного лица']]
    if not all(mandatory):
        st.error("⚠️ Заполните обязательные поля (отмечены *)!")
    else:
        with st.spinner("Создаем шедевр и отправляем коллегам..."):
            f_score = min(score, 100)
            report_bytes = make_expert_excel(client_info, data, f_score)
            
            try:
                url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
                caption = (f"🚀 *Коллеги, у нас новый заказ. Давайте зарабатывать!*\n\n"
                           f"🏢 *Компания:* {client_info['Наименование компании']}\n"
                           f"📊 *Зрелость ИТ:* {f_score}%\n"
                           f"👤 *Эксперт:* Иван Рудой\n"
                           f"📞 *Тел:* {client_info['Контактный телефон']}")
                
                files = {'document': (f"Audit_{client_info['Наименование компании']}_2026.xlsx", report_bytes)}
                r = requests.post(url, data={"chat_id": CHAT_ID, "caption": caption, "parse_mode": "Markdown"}, files=files)
                
                if r.ok:
                    st.success("Отчет готов и отправлен!")
                    st.balloons()
            except Exception as e:
                st.error(f"Сбой: {e}")
            
            st.download_button(f"📥 Скачать отчет для {client_info['Наименование компании']}", report_bytes, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v2.0 | Almaty 2026")
