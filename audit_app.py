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

# --- НАСТРОЙКИ TELEGRAM ---
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ЛОГОТИП И КОНТАКТЫ ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### Мы поможем Вам стать лучше!")
st.divider()

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v6.3")

data = {}
client_info = {}
validation_errors = []
score = 0

# --- ШАПКА: ОБЩАЯ ИНФОРМАЦИЯ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город:*")
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    site_input = st.text_input("Сайт компании:*", placeholder="example.kz")
    client_info['Сайт компании'] = site_input
    custom_email_mode = st.checkbox("Email отличается от сайта")
    
    if custom_email_mode:
        client_info['Email'] = st.text_input("Email контактного лица:*")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица (логин):*")
            e_col1, e_col2 = st.columns([2, 3])
            with e_col1:
                email_prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="email_pre")
            with e_col2:
                st.markdown(f"<div style='padding-top: 5px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
            client_info['Email'] = f"{email_prefix}@{clean_domain}" if email_prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    st.write("Контактный телефон:*")
    p_col1, p_col2 = st.columns([1, 2])
    country_codes = [("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998")]
    selected_code = p_col1.selectbox("Код", country_codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed")
    client_info['Контактный телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

st.divider()

# --- БЛОК 1: ИТ ---
st.header("Блок 1: Информационные технологии")

# 1.1 АРМ
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)
data['1.1. Всего АРМ'] = total_arm
selected_os_arm = st.multiselect("ОС на АРМ:", ["Windows XP/Vista/7/8", "Windows 10", "Windows 11", "Linux", "macOS", "Другое"])
sum_os_arm = 0
for os_item in selected_os_arm:
    val = st.number_input(f"Кол-во на {os_item}:", min_value=0, step=1, key=f"arm_{os_item}")
    data[f"ОС АРМ ({os_item})"] = val
    sum_os_arm += val
if total_arm > 0 and sum_os_arm != total_arm:
    st.warning(f"⚠️ Ошибка: Всего АРМ {total_arm}, по ОС {sum_os_arm}.")
    validation_errors.append("Несовпадение АРМ")

# 1.2 Сетевая инфраструктура
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_toggle"):
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL/VDSL", "Нет"]
    col_net1, col_net2 = st.columns(2)
    with col_net1:
        st.write("**Каналы связи:**")
        main_type = st.selectbox("Основной канал:", net_types, key="main_net_type")
        main_speed = st.number_input("Скорость (Mbit/s):", min_value=0, step=10, key="main_net_speed")
        data['1.2.1. Основной канал'] = f"{main_type} ({main_speed} Mbit/s)"
        data['main_speed_val'] = main_speed
    with col_net2:
        st.write("**Резервный канал:**")
        back_type = st.selectbox("Резервный канал:", net_types, key="back_net_type")
        back_speed = st.number_input("Скорость резервного (Mbit/s):", min_value=0, step=10, key="back_net_speed")
        data['1.2.2. Резервный канал'] = f"{back_type} ({back_speed} Mbit/s)"

    st.write("**Активное сетевое оборудование:**")
    c_net1, c_net2, c_net3 = st.columns(3)
    with c_net1:
        if st.checkbox("Маршрутизаторы"):
            data['1.2.4. Маршрутизаторы'] = st.number_input("Кол-во маршрутизаторов:", min_value=0, key="n_rout")
    with c_net2:
        if st.checkbox("Коммутаторы L2"):
            data['1.2.5. Коммутаторы L2'] = st.number_input("Кол-во L2:", min_value=0, key="n_l2")
    with c_net3:
        if st.checkbox("Коммутаторы L3"):
            data['1.2.6. Коммутаторы L3'] = st.number_input("Кол-во L3:", min_value=0, key="n_l3")

    st.write("**Уровни сети:**")
    l_col1, l_col2, l_col3 = st.columns(3)
    with l_col1:
        if st.checkbox("Ядро (Core)"): data['Уровень сети: Ядро'] = st.text_input("Производитель Core:")
    with l_col2:
        if st.checkbox("Распределение"): data['Уровень сети: Распределение'] = st.text_input("Производитель Dist:")
    with l_col3:
        if st.checkbox("Доступ"): data['Уровень сети: Доступ'] = st.text_input("Производитель Access:")

    if st.checkbox("Wi-Fi"):
        w_col1, w_col2, w_col3 = st.columns(3)
        with w_col1:
            if st.checkbox("Контроллер Wi-Fi"): data['Wi-Fi Контроллер'] = st.text_input("Модель контроллера:")
        with w_col2:
            data['Wi-Fi AP'] = st.number_input("Кол-во точек доступа:", min_value=0, key="n_ap")
        with w_col3:
            data['Wi-Fi Стандарт'] = st.selectbox("Стандарт:", ["Wi-Fi 6/6E", "Wi-Fi 5", "Другое"])

    if st.checkbox("Межсетевой экран (NGFW)", key="ngfw_chk"):
        ngfw_vendor = st.text_input("Производитель NGFW:")
        data['1.2.7. NGFW'] = f"Да ({ngfw_vendor})"
        score += 20

# 1.3 Серверы
st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
col_s1, col_s2 = st.columns(2)
with col_s1:
    phys_count = st.number_input("Физические серверы:", min_value=0, step=1)
    data['1.3.1. Физические серверы'] = phys_count
with col_s2:
    virt_count = st.number_input("Виртуальные серверы:", min_value=0, step=1)
    data['1.3.2. Виртуальные серверы'] = virt_count

s_os_list = ["Windows Server 2008/2012 R2", "Windows Server 2016", "Windows Server 2019", "Windows Server 2022", "Linux", "Другое"]
selected_os_srv = st.multiselect("ОС серверов:", s_os_list)
for os_s in selected_os_srv:
    data[f"ОС Сервера ({os_s})"] = st.number_input(f"Кол-во на {os_s}:", min_value=0, key=f"srv_{os_s}")

# 1.4 СХД (ПОЛНОЕ ВОССТАНОВЛЕНИЕ + ALL-FLASH)
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Наличие СХД", key="storage_toggle"):
    st_col1, st_col2 = st.columns(2)
    with st_col1:
        st_arch = st.selectbox("Архитектура массива:", ["All-Flash (NVMe/SSD)", "Hybrid (SSD+HDD)", "HDD Only"])
        st_type = st.selectbox("Тип подключения:", ["SAN (FC/iSCSI)", "NAS (NFS/SMB)", "DAS (Direct)", "Облачное"])
        st_vendor = st.text_input("Производитель СХД:")
    with st_col2:
        st_cap = st.number_input("Общая полезная емкость (ТБ):", min_value=0)
        st_prot = st.multiselect("Используемые протоколы:", ["FC (Fibre Channel)", "iSCSI", "NFS", "SMB/CIFS", "NVMe-oF", "S3"])
    
    data['1.4. СХД'] = f"{st_vendor} | {st_arch} | {st_type} ({st_cap} TB)"
else:
    data['1.4. СХД'] = "Нет"

if st.checkbox("Резервное копирование"):
    v_n_b = st.text_input("Производитель системы резервного копирования:")
    data["Резервное копирование"] = f"Да ({v_n_b})"
    score += 20

# 1.5 ИС
st.write("---")
st.subheader("1.5. Внутренние Информационные системы")
if st.toggle("ИС организации", key="is_toggle"):
    m_sys = st.selectbox("Почта:", ["Exchange (On-Prem)", "Lotus", "Microsoft 365", "Google Workspace", "Нет"])
    data['mail_type'] = m_sys
    if m_sys in ["Exchange (On-Prem)", "Lotus"]:
        m_ver = st.text_input(f"Версия {m_sys}:")
        data['1.5.1. Почтовая система'] = f"{m_sys} (v.{m_ver})"
    else:
        data['1.5.1. Почтовая система'] = m_sys
    for is_name in ["1С", "Битрикс24", "Documentolog"]:
        if st.checkbox(is_name): data[f"ИС: {is_name}"] = st.text_input(f"Версия {is_name}:")

st.divider()

# --- БЛОК 2: ИБ ---
st.header("Блок 2: Информационная Безопасность")
if st.toggle("Средства защиты", key="ib_toggle"):
    ib_systems = {
        "EPP (Антивирус)": 10, "DLP (Утечки)": 15, "PAM (Привилегии)": 10,
        "SIEM (Мониторинг)": 20, "VM (Уязвимости)": 10, "EDR/XDR (Точки)": 15,
        "WAF (Веб)": 10, "Sandbox (Песочница)": 5, "IDS/IPS (Атаки)": 5, "IDM/IGA (Доступ)": 5,
        "MFA (Аутентификация)": 15, "Anti-DDoS": 15
    }
    col1, col2 = st.columns(2)
    items = list(ib_systems.items())
    for i, (label, pts) in enumerate(items):
        with (col1 if i < 6 else col2):
            if st.checkbox(label):
                v_n = st.text_input(f"Производитель {label}:", key=f"vn_{label}")
                data[label] = f"Да ({v_n if v_n else 'не указан'})"
                score += pts
            else:
                data[label] = "Нет"

st.divider()

# --- БЛОК 3: WEB ---
st.header("Блок 3: Web-ресурсы")
web_active = st.toggle("Наличие Web", key="w_t")
if web_active:
    data['3.1. Хостинг'] = st.selectbox("Хостинг:", ["Собственный ЦОД", "Облако KZ", "Облако Global"])
    data['3.2. Frontend'] = st.multiselect("Frontend:", ["Nginx", "Apache", "IIS"])

# --- БЛОК 4: РАЗРАБОТКА ---
st.header("Блок 4: Разработка")
dev_active = st.toggle("Наличие Разработки", key="d_t")
if dev_active:
    data['4.1. Разработчики'] = st.number_input("Количество разработчиков:", min_value=0)
    data['4.3. CI/CD'] = st.checkbox("Используется CI/CD")
    data['4.2. Стек'] = st.text_input("Стек технологий:")

# --- ГЕНЕРАЦИЯ EXCEL (УМНАЯ ЛОГИКА) ---
def make_expert_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Report"
    
    h_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    w_font = Font(color="FFFFFF", bold=True)
    brd = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    ws.merge_cells('A1:D2'); ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ ПО ИТ И ИБ (2026)"; ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

    curr_row = 4
    for k, v in c_info.items():
        ws.cell(row=curr_row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=curr_row, column=2, value=str(v)); curr_row += 1
    
    ws.cell(row=curr_row, column=1, value="ИНДЕКС ЗРЕЛОСТИ:").font = Font(bold=True)
    ws.cell(row=curr_row, column=2, value=f"{final_score}%"); curr_row += 3

    headers = ["Параметр", "Значение", "Статус", "Рекомендация / Обоснование"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=curr_row, column=i, value=h)
        c.fill = h_fill; c.font = w_font; c.border = brd
    
    curr_row += 1
    
    n_arm = results.get('1.1. Всего АРМ', 0)
    n_srv = results.get('1.3.1. Физические серверы', 0) + results.get('1.3.2. Виртуальные серверы', 0)
    has_edr = results.get('EDR/XDR (Точки)') != "Нет"
    has_web = web_active or dev_active
    
    for k, v in results.items():
        if any(x in k for x in ["ОС АРМ", "ОС Сервера", "speed", "mail_type"]): continue
        ws.cell(row=curr_row, column=1, value=k).border = brd
        ws.cell(row=curr_row, column=2, value=str(v)).border = brd
        
        status, rec, color = "В норме", "Поддерживать состояние.", "000000"

        if v == "Нет":
            if k == "1.4. СХД":
                if n_srv > 15: status, rec, color = "КРИТИЧНО", "При 15+ серверах отсутствие выделенной СХД критически снижает надежность.", "FF0000"
                else: status, rec, color = "ВНИМАНИЕ", "Рекомендуется к приобретению для централизации данных.", "FFC000"
            elif k == "SIEM (Мониторинг)":
                if n_arm < 100 and n_srv < 20 and not has_edr:
                    status, rec, color = "ВНИМАНИЕ", "Малый масштаб и отсутствие EDR. SIEM нецелесообразен на данном этапе.", "FFC000"
                else:
                    status, rec, color = "КРИТИЧНО", "Необходим автоматизированный сбор и корреляция событий.", "FF0000"
            elif k == "EDR/XDR (Точки)":
                if n_arm < 50: status, rec, color = "РЕКОМЕНДУЕТСЯ К ПРИОБРЕТЕНИЮ", "Для защиты от Ransomware и сложных угроз.", "00B050"
                else: status, rec, color = "КРИТИЧНО", "Классический антивирус неэффективен при текущем масштабе.", "FF0000"
            elif k == "MFA (Аутентификация)":
                if n_arm < 20: status, rec, color = "ВНИМАНИЕ", "Рекомендуется к приобретению для защиты удаленного доступа.", "FFC000"
                else: status, rec, color = "КРИТИЧНО", "Второй фактор обязателен для защиты учетных записей.", "FF0000"
            elif k == "WAF (Веб)":
                if not has_web: status, rec, color = "НЕ ТРЕБУЕТСЯ", "Внешние веб-сервисы не обнаружены.", "000000"
                else: status, rec, color = "КРИТИЧНО", "Веб-ресурсы открыты для SQL-инъекций и XSS атак.", "FF0000"
            else:
                status, rec, color = "РИСК", "Рекомендуется к приобретению для усиления ИБ.", "FF0000"

        st_c = ws.cell(row=curr_row, column=3, value=status); st_c.font = Font(color=color, bold=True); st_c.border = brd
        ws.cell(row=curr_row, column=4, value=rec).border = brd
        curr_row += 1

    for c, w in {'A': 30, 'B': 25, 'C': 15, 'D': 50}.items(): ws.column_dimensions[c].width = w
    wb.save(output); return output.getvalue()

# --- ФИНАЛ ---
st.divider()
if st.button("📊 Сформировать экспертный отчет", disabled=len(validation_errors) > 0):
    if not all([client_info['Город'], client_info['Наименование компании'], client_info['Email']]):
        st.error("⚠️ Заполните все поля со звездочкой!")
    else:
        with st.spinner("Генерация..."):
            f_score = min(score, 100); r_bytes = make_expert_excel(client_info, data, f_score)
            try:
                requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", data={"chat_id": CHAT_ID, "caption": f"Аудит: {client_info['Наименование компании']}"}, files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", r_bytes)})
            except: pass
            st.success("Отчет готов!")
            st.download_button("📥 Скачать Excel", r_bytes, f"Audit_{client_info['Наименование компании']}.xlsx")
