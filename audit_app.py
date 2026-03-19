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

st.markdown("### Мы поможем Вам стать лучше!")
st.divider()

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v2.3")

data = {}
client_info = {}
score = 0

# --- ШАПКА: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город:*")
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    
    # Сайт компании
    site_input = st.text_input("Сайт компании:*", key="site_field", placeholder="example.kz")
    client_info['Сайт компании'] = site_input

    # ЛОГИКА EMAIL С ЧЕКБОКСОМ
    custom_email_mode = st.checkbox("Email отличается от сайта (не рекомендуется)")
    
    if custom_email_mode:
        client_info['Email'] = st.text_input("Email контактного лица:*", placeholder="info@other-domain.com")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица (только логин до @):*")
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
    
    # ВВОД НОМЕРА ТЕЛЕФОНА С ПРЕФИКСОМ
    st.write("Контактный телефон:*")
    p_col1, p_col2 = st.columns([1, 2])
    country_codes = [
        ("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"),
        ("🇰🇬 +996", "+996"), ("🇹🇯 +992", "+992"), ("🇦🇪 +971", "+971"),
        ("🇹🇷 +90", "+90"), ("🇦🇿 +994", "+994"), ("🇧🇾 +375", "+375"),
        ("🇬🇪 +995", "+995"), ("🇺🇸 +1", "+1"), ("🇬🇧 +44", "+44")
    ]
    selected_code = p_col1.selectbox("Код страны", country_codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер телефона", placeholder="777 777 77 77", label_visibility="collapsed")
    client_info['Контактный телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

st.divider()

# --- БЛОК 1: ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ ---
st.header("Блок 1: Информационные технологии")

# 1.1 Конечные точки
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
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink"]
    data['1.2.1. Основной канал'] = st.selectbox("Тип канала:", net_types, key="net_type")
    data['1.2.2. NGFW'] = st.text_input("Вендор Межсетевого экрана (NGFW):", key="ngfw_v")
    if data['1.2.2. NGFW']: score += 20
else:
    data['1.2. Сетевая инфраструктура'] = "Не указана/Аренда"

# 1.3 Серверы и Виртуализация
st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
col_s1, col_s2 = st.columns(2)
with col_s1:
    phys_servers = st.number_input("Количество физических серверов:", min_value=0, step=1, key="phys_srv")
    data['1.3.1. Физические серверы'] = phys_servers
with col_s2:
    virt_servers = st.number_input("Количество виртуальных серверов:", min_value=0, step=1, key="virt_srv")
    data['1.3.2. Виртуальные серверы'] = virt_servers

selected_os_srv = st.multiselect("Выберите ОС серверов:", ["Windows Server", "Linux", "Unix", "Другое"], key="ms_srv_list")
if selected_os_srv:
    for os_s in selected_os_srv:
        count_srv = st.number_input(f"Количество серверов на {os_s}:", min_value=0, step=1, key=f"srv_cnt_{os_s}")
        data[f"ОС Сервера ({os_s})"] = count_srv

selected_virt_sys = st.multiselect("Выберите системы виртуализации:", ["VMware", "Hyper-V", "Proxmox", "KVM", "Другое", "Нет"], key="virt_sys_list")
if selected_virt_sys:
    if "Нет" in selected_virt_sys:
        data['1.3.3. Виртуализация'] = "Нет"
    else:
        for v_sys in selected_virt_sys:
            v_cnt = st.number_input(f"Количество хостов {v_sys}:", min_value=0, step=1, key=f"v_cnt_{v_sys}")
            data[f"Система виртуализации ({v_sys})"] = v_cnt

# 1.4 Системы хранения данных (СХД) - НОВЫЙ ПУНКТ
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Есть собственная СХД", key="storage_toggle"):
    data['1.4.1. Типы носителей'] = st.multiselect("Типы носителей:", ["HDD (NL-SAS / SATA)", "SSD (SATA / SAS)", "NVMe", "SCM (Storage Class Memory)"], key="st_media")
    
    col_pct1, col_pct2 = st.columns(2)
    with col_pct1:
        data['1.4.2. Доля HDD (%)'] = st.number_input("Процент HDD:", min_value=0, max_value=100, step=5, key="pct_hdd")
    with col_pct2:
        data['1.4.3. Доля SSD (%)'] = st.number_input("Процент SSD:", min_value=0, max_value=100, step=5, key="pct_ssd")
    
    col_chk1, col_chk2 = st.columns(2)
    data['1.4.4. Гибридная СХД'] = col_chk1.checkbox("Используется гибридная СХД (Hybrid Storage)", key="hybrid_st")
    data['1.4.5. All-Flash'] = col_chk2.checkbox("Есть All-Flash массивы", key="allflash_st")
    
    data['1.4.6. RAID-группы'] = st.multiselect("Используемые RAID-группы:", ["RAID 0", "RAID 1", "RAID 5", "RAID 6", "RAID 10", "RAID 50", "RAID 60", "JBOD"], key="raid_list")
else:
    data['1.4. СХД'] = "Не используется/Облако"

# 1.5 Почта и мониторинг
st.write("---")
col_p1, col_p2 = st.columns(2)
with col_p1:
    st.subheader("1.5. Почтовая система")
    data['1.5. Почта'] = st.selectbox("Тип почты:", ["Exchange (On-Prem)", "Microsoft 365", "Google Workspace", "Yandex/Mail.ru Cloud", "Собственный сервер", "Нет"], key="mail_sys")
with col_p2:
    st.subheader("1.6. Система мониторинга")
    has_mon = st.checkbox("Есть ли система мониторинга?", key="mon_chk")
    data['1.6. Мониторинг'] = st.selectbox("Система:", ["Zabbix", "Nagios", "PRTG", "Prometheus", "Другое"], key="mon_sel") if has_mon else "Нет"

st.subheader("1.7. Внутренние Информационные системы")
has_is = st.checkbox("Есть ли внутренние ИС (1C, ERP, CRM)?", key="is_chk")
data['1.7. Внутренние ИС'] = st.text_input("Перечислите:", key="is_input") if has_is else "Нет"

st.divider()

# --- БЛОК 2: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ ---
st.header("Блок 2: Информационная Безопасность")
if st.toggle("Есть отдел ИБ", key="ib_toggle"):
    ib_list = {"DLP": 15, "PAM": 10, "SIEM": 20, "WAF": 10, "EDR": 15, "Резервное копирование": 20}
    for label, pts in ib_list.items():
        if st.checkbox(label, key=f"ib_{label}"):
            v_n = st.text_input(f"Вендор {label}:", key=f"vn_{label}")
            data[label] = f"Да ({v_n if v_n else 'не указан'})"
            score += pts
        else:
            data[label] = "Нет"

# --- БЛОК 3: WEB-РЕСУРСЫ ---
st.header("Блок 3: Web-ресурсы")
if st.toggle("Есть свои Web-ресурсы", key="web_toggle"):
    data['3.1. Хостинг'] = st.selectbox("Хостинг:", ["Собственный ЦОД", "Облако (KZ)", "Облако (Global)"], key="host")
    data['3.2. Frontend'] = st.multiselect("Frontend серверы:", ["Nginx", "Apache", "IIS", "LiteSpeed", "Caddy", "Cloudflare"], key="fnt")

# --- БЛОК 4: РАЗРАБОТКА ---
st.header("Блок 4: Разработка")
if st.toggle("Своя разработка", key="dev_toggle"):
    data['4.1. Разработчики'] = st.number_input("Кол-во разработчиков:", min_value=0, key="dev_c")
    data['4.2. CI/CD'] = st.checkbox("CI/CD используется", key="cicd_c")

# --- ГЕНЕРАЦИЯ EXCEL ---
def make_expert_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Khalil Audit Report"

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    ws.merge_cells('A1:D2')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ ПО ИТ И ИБ (2026)"
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=16, color="1F4E78")

    current_row = 4
    for k, v in c_info.items():
        ws.cell(row=current_row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=current_row, column=2, value=str(v))
        current_row += 1
    
    auto_date = datetime.now().strftime("%d.%m.%Y %H:%M")
    ws.cell(row=current_row, column=1, value="Дата генерации отчета:").font = Font(bold=True)
    ws.cell(row=current_row, column=2, value=auto_date)
    current_row += 2

    ws.cell(row=current_row, column=1, value="ИНДЕКС ТЕХНИЧЕСКОЙ ЗРЕЛОСТИ:").font = Font(bold=True)
    score_cell = ws.cell(row=current_row, column=2, value=f"{final_score}%")
    bg_color = "92D050" if final_score > 70 else "FFC000" if final_score > 40 else "FF7C80"
    score_cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    score_cell.font = Font(bold=True)
    current_row += 2

    headers = ["Параметр", "Значение", "Статус", "Рекомендация эксперта Khalil Trade"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=current_row, column=i, value=h)
        cell.fill = header_fill; cell.font = white_font

    current_row += 1
    rec_map = {
        "Нет": "Требуется внедрение для минимизации рисков.", 
        "Резервное копирование": "Критично! Настроить схему 3-2-1.", 
        "NGFW": "Рекомендуется для защиты периметра."
    }

    for k, v in results.items():
        ws.cell(row=current_row, column=1, value=k).border = border
        ws.cell(row=current_row, column=2, value=str(v)).border = border
        status = "В норме"; recommendation = "Поддерживать текущее состояние."
        
        if "Нет" in str(v) or v == 0 or v == []:
            status = "РИСК"
            recommendation = rec_map.get(k, "Рассмотреть возможность внедрения.")
            st_cell = ws.cell(row=current_row, column=3, value=status)
            st_cell.font = Font(color="FF0000", bold=True)
        else:
            ws.cell(row=current_row, column=3, value=status)
        
        ws.cell(row=current_row, column=4, value=recommendation).border = border
        ws.cell(row=current_row, column=3).border = border
        current_row += 1

    for col, width in {'A': 35, 'B': 30, 'C': 15, 'D': 60}.items():
        ws.column_dimensions[col].width = width
    
    wb.save(output)
    return output.getvalue(), auto_date

# --- ФИНАЛ И ОТПРАВКА ---
st.divider()
if st.button("📊 Сформировать экспертный отчет", key="btn_final"):
    mandatory = [
        client_info['Город'], 
        client_info['Наименование компании'], 
        client_info['ФИО контактного лица'], 
        client_info['Сайт компании'], 
        client_info.get('Email'),
        client_info.get('Контактный телефон')
    ]
    if not all(mandatory):
        st.error("⚠️ Заполните все обязательные поля (включая Сайт, Email и Телефон)!")
    else:
        with st.spinner("Создаем отчет..."):
            f_score = min(score, 100)
            report_bytes, final_date = make_expert_excel(client_info, data, f_score)
            try:
                url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
                caption = (f"🚀 *Новый заказ Khalil Trade*\n\n"
                           f"🏢 *Компания:* {client_info['Наименование компании']}\n"
                           f"📊 *Зрелость ИТ:* {f_score}%\n"
                           f"📅 *Дата:* {final_date}\n"
                           f"👤 *Контакт:* {client_info['ФИО контактного лица']}\n"
                           f"📧 *Email:* {client_info['Email']}\n"
                           f"📞 *Тел:* {client_info['Контактный телефон']}")
                
                files = {'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report_bytes)}
                requests.post(url, data={"chat_id": CHAT_ID, "caption": caption, "parse_mode": "Markdown"}, files=files)
                st.success("Отчет успешно отправлен в Telegram!")
                st.balloons()
            except Exception as e:
                st.error(f"Ошибка связи: {e}")
            st.download_button(f"📥 Скачать отчет", report_bytes, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v2.3 | Almaty 2026")
