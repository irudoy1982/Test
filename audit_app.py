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
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    
    # Сайт компании
    site_input = st.text_input("Сайт компании:*", key="site_field", placeholder="example.kz")
    client_info['Сайт компании'] = site_input

    # ЛОГИКА EMAIL С ЧЕКБОКСОМ
    custom_email_mode = st.checkbox("Email отличается от сайта")
    
    if custom_email_mode:
        client_info['Email'] = st.text_input("Email контактного лица:*", placeholder="info@other-domain.com")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица (логин):*")
            e_col1, e_col2 = st.columns([2, 3])
            with e_col1:
                email_prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="email_pre")
            with e_col2:
                st.markdown(f"<div style='padding-top: 5px; font-size: 16px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
            client_info['Email'] = f"{email_prefix}@{clean_domain}" if email_prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    
    # --- ПОЛЕ ТЕЛЕФОНА (Версия 2.0 - Native Streamlit) ---
    st.write("Контактный телефон:*")
    p_col1, p_col2 = st.columns([1, 2])
    with p_col1:
        # Список популярных префиксов
        country_data = st.selectbox(
            "Код",
            options=[
                ("🇰🇿 +7", "+7"), 
                ("🇷🇺 +7", "+7"), 
                ("🇺🇿 +998", "+998"), 
                ("🇰🇬 +996", "+996"), 
                ("🇦🇪 +971", "+971"),
                ("🇹🇷 +90", "+90")
            ],
            format_func=lambda x: x[0],
            label_visibility="collapsed"
        )
    with p_col2:
        phone_raw = st.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed")
    
    # Сборка полного номера
    client_info['Контактный телефон'] = f"{country_data[1]} {phone_raw}"

st.divider()

# --- БЛОК 1: ИТ ИНФРАСТРУКТУРА ---
st.header("Блок 1: Информационные технологии")

# 1.1 АРМ
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1, key="total_arm_val")
data['1.1. Всего АРМ'] = total_arm
selected_os_arm = st.multiselect("Выберите ОС на АРМ:", ["Windows", "Linux", "macOS", "Другое"], key="ms_arm_list")
if selected_os_arm:
    for os_item in selected_os_arm:
        count_arm = st.number_input(f"Количество АРМ на {os_item}:", min_value=0, step=1, key=f"arm_cnt_{os_item}")
        data[f"ОС АРМ ({os_item})"] = count_arm

# 1.2 Сетевая инфраструктура
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_toggle"):
    data['1.2.1. Основной канал'] = st.selectbox("Тип канала:", ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink"])
    data['1.2.2. NGFW'] = st.text_input("Вендор NGFW (Fortinet, Checkpoint и т.д.):")
    if data['1.2.2. NGFW']: score += 20
else:
    data['1.2. Сетевая инфраструктура'] = "Аренда/Нет"

# 1.3 Серверы
st.write("---")
st.subheader("1.3. Серверы")
col_s1, col_s2 = st.columns(2)
with col_s1:
    data['1.3. Физические серверы'] = st.number_input("Кол-во физических серверов:", min_value=0, step=1)
with col_s2:
    data['1.3. Виртуальные серверы'] = st.number_input("Кол-во виртуальных серверов:", min_value=0, step=1)

# 1.4-1.7
col_v1, col_v2 = st.columns(2)
with col_v1:
    data['1.4. Виртуализация'] = st.multiselect("Системы виртуализации:", ["VMware", "Hyper-V", "Proxmox", "KVM", "Нет"])
with col_v2:
    data['1.5. Почтовая система'] = st.selectbox("Почта:", ["Exchange", "M365", "Google", "Yandex", "Свой сервер", "Нет"])

data['1.6. Внутренние ИС'] = st.text_input("ИС (1C, ERP, CRM):", placeholder="Перечислите через запятую")
data['1.7. Мониторинг'] = st.selectbox("Система мониторинга:", ["Нет", "Zabbix", "Nagios", "PRTG", "Prometheus"])

st.divider()

# Блок 2: ИБ
st.header("Блок 2: Информационная Безопасность")
if st.toggle("Есть выделенные системы ИБ", key="ib_toggle"):
    ib_list = {"DLP": 15, "PAM": 10, "SIEM": 20, "WAF": 10, "EDR": 15, "Резервное копирование": 20}
    for label, pts in ib_list.items():
        if st.checkbox(label, key=f"ib_{label}"):
            v_n = st.text_input(f"Вендор {label}:", key=f"vn_{label}")
            data[label] = f"Да ({v_n if v_n else 'не указан'})"
            score += pts
        else:
            data[label] = "Нет"

st.divider()

# --- ЭКСЕЛЬ ГЕНЕРАТОР ---
def make_expert_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Khalil Audit Report"
    
    # Стили
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Шапка
    ws.merge_cells('A1:D2')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ ПО ИТ И ИБ (2026)"
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=16, color="1F4E78")

    # Информация о клиенте
    curr = 4
    for k, v in c_info.items():
        ws.cell(row=curr, column=1, value=k).font = Font(bold=True)
        ws.cell(row=curr, column=2, value=str(v))
        curr += 1
    
    auto_date = datetime.now().strftime("%d.%m.%Y %H:%M")
    ws.cell(row=curr, column=1, value="Дата:").font = Font(bold=True)
    ws.cell(row=curr, column=2, value=auto_date)
    
    curr += 2
    ws.cell(row=curr, column=1, value="ИНДЕКС ТЕХНИЧЕСКОЙ ЗРЕЛОСТИ:").font = Font(bold=True)
    sc_cell = ws.cell(row=curr, column=2, value=f"{final_score}%")
    sc_cell.font = Font(bold=True)
    
    curr += 2
    h_labels = ["Параметр", "Значение", "Статус", "Рекомендация эксперта"]
    for i, h in enumerate(h_labels, 1):
        c = ws.cell(row=curr, column=i, value=h)
        c.fill = header_fill; c.font = white_font
    
    curr += 1
    for k, v in results.items():
        ws.cell(row=curr, column=1, value=k).border = border
        ws.cell(row=curr, column=2, value=str(v)).border = border
        ws.cell(row=curr, column=3, value="Проверено").border = border
        ws.cell(row=curr, column=4, value="Поддерживать актуальность").border = border
        curr += 1

    for col, width in {'A': 35, 'B': 30, 'C': 15, 'D': 50}.items():
        ws.column_dimensions[col].width = width
        
    wb.save(output)
    return output.getvalue(), auto_date

# --- ФИНАЛ ---
if st.button("📊 Сформировать экспертный отчет", key="btn_final"):
    # Проверка обязательных полей
    is_valid_phone = len(phone_raw.strip()) > 5
    mandatory = [client_info['Город'], client_info['Наименование компании'], client_info['ФИО контактного лица'], client_info['Email']]
    
    if not all(mandatory) or not is_valid_phone:
        st.error("⚠️ Пожалуйста, заполните все поля со звездочкой (*), включая телефон и почту!")
    else:
        with st.spinner("Отправка данных..."):
            f_score = min(score, 100)
            excel_bytes, final_date = make_expert_excel(client_info, data, f_score)
            
            try:
                caption = (f"🚀 *Новый аудит Khalil Trade*\n\n"
                           f"🏢 *Компания:* {client_info['Наименование компании']}\n"
                           f"📊 *Зрелость:* {f_score}%\n"
                           f"👤 *Контакт:* {client_info['ФИО контактного лица']}\n"
                           f"📞 *Тел:* {client_info['Контактный телефон']}\n"
                           f"📧 *Email:* {client_info['Email']}")
                
                requests.post(
                    f"https://api.telegram.org/bot{TOKEN}/sendDocument",
                    data={"chat_id": CHAT_ID, "caption": caption, "parse_mode": "Markdown"},
                    files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", excel_bytes)}
                )
                st.success("Отчет успешно отправлен в Telegram!")
                st.balloons()
                st.download_button("📥 Скачать Excel копию", excel_bytes, f"Audit_{client_info['Наименование компании']}.xlsx")
            except Exception as e:
                st.error(f"Ошибка связи: {e}")

st.info("Khalil Audit System v2.0 | Almaty 2026")
