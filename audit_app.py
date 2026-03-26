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

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v4.1")

# --- РАСШИРЕННАЯ ИНСТРУКЦИЯ ДЛЯ ПОЛЬЗОВАТЕЛЯ ---
with st.expander("📖 Инструкция по заполнению опросника (нажмите, чтобы развернуть)"):
    st.markdown("""
    ### Добро пожаловать в систему технического аудита!
    Этот инструмент поможет провести экспресс-анализ вашей ИТ-инфраструктуры и информационной безопасности.

    **Рекомендации по заполнению:**
    1. **Подготовьте данные:** Для точности отчета желательно иметь под рукой информацию о количестве серверов, параметрах интернет-каналов и версиях ключевых бизнес-систем (1С, ERP и др.).
    2. **Точность полей:** Поля, отмеченные звездочкой (**\***), обязательны для заполнения. Без них формирование финального файла будет невозможно.
    3. **Блок ИТ-инфраструктуры:** Указывайте реальные скорости каналов и типы носителей в СХД — это влияет на экспертные рекомендации в отчете.
    4. **Информационные системы:** При выборе систем (например, 1С или SAP) обязательно указывайте их версию или конфигурацию в появившемся поле ввода.
    5. **Безопасность:** Отметьте те системы защиты, которые уже внедрены. Если системы нет — оставьте поле пустым, система автоматически добавит рекомендацию по её внедрению.

    **Как получить результат?**
    После заполнения всех блоков нажмите кнопку **«Сформировать экспертный отчет»** внизу страницы. Вы сможете скачать готовый файл Excel с анализом рисков и рекомендациями экспертов Khalil Trade.
    """)

data = {}
client_info = {}
score = 0

# --- ШАПКА: ИНФОРМАЦИЯ О КЛИЕНТЕ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город:*")
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    
    site_input = st.text_input("Сайт компании:*", key="site_field", placeholder="example.kz")
    client_info['Сайт компании'] = site_input

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
    
    st.write("Контактный телефон:*")
    p_col1, p_col2 = st.columns([1, 2])
    country_codes = [
        ("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"),
        ("🇰🇬 +996", "+996"), ("🇹🇯 +992", "+992"), ("🇦🇪 +971", "+971")
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
selected_os_arm = st.multiselect("Выберите ОС на АРМ:", ["Windows XP/Vista/7/8", "Windows 10", "Windows 11", "Linux", "macOS", "Другое"], key="ms_arm_list")
if selected_os_arm:
    for os_item in selected_os_arm:
        count_arm = st.number_input(f"Количество АРМ на {os_item}:", min_value=0, step=1, key=f"arm_cnt_{os_item}")
        data[f"ОС АРМ ({os_item})"] = count_arm

# 1.2 Сетевая инфраструктура
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_toggle"):
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL/VDSL", "Нет"]
    col_net1, col_net2 = st.columns(2)
    with col_net1:
        st.write("**Основной канал:**")
        main_type = st.selectbox("Тип (основной):", net_types, key="main_net_type")
        main_speed = st.number_input("Скорость основного (Mbit/s):", min_value=0, step=10, key="main_net_speed")
        data['1.2.1. Основной канал'] = f"{main_type} ({main_speed} Mbit/s)"
    with col_net2:
        st.write("**Резервный канал:**")
        back_type = st.selectbox("Тип (резервный):", net_types, key="back_net_type")
        back_speed = st.number_input("Скорость резервного (Mbit/s):", min_value=0, step=10, key="back_net_speed")
        data['1.2.2. Резервный канал'] = f"{back_type} ({back_speed} Mbit/s)"

    st.write("**Активное сетевое оборудование:**")
    c_net1, c_net2, c_net3 = st.columns(3)
    with c_net1:
        if st.checkbox("Маршрутизаторы", key="router_chk"):
            r_count = st.number_input("Кол-во маршрутизаторов:", min_value=0, step=1, key="router_cnt")
            data['1.2.4. Маршрутизаторы'] = f"Да ({r_count} шт)"
    with c_net2:
        if st.checkbox("Коммутаторы L2", key="swl2_chk"):
            sw2_count = st.number_input("Кол-во коммутаторов L2:", min_value=0, step=1, key="swl2_cnt")
            data['1.2.5. Коммутаторы L2'] = f"Да ({sw2_count} шт)"
    with c_net3:
        if st.checkbox("Коммутаторы L3", key="swl3_chk"):
            sw3_count = st.number_input("Кол-во коммутаторов L3:", min_value=0, step=1, key="swl3_cnt")
            data['1.2.6. Коммутаторы L3'] = f"Да ({sw3_count} шт)"

    st.write("**Уровни сети:**")
    l_col1, l_col2, l_col3 = st.columns(3)
    with l_col1:
        if st.checkbox("Ядро (Core)", key="net_core"):
            core_v = st.text_input("Основной производитель (Core):", key="core_vendor")
            data['Уровень сети: Ядро'] = core_v if core_v else "Да"
    with l_col2:
        if st.checkbox("Уровень распределения", key="net_dist"):
            dist_v = st.text_input("Основной производитель (Dist):", key="dist_vendor")
            data['Уровень сети: Распределение'] = dist_v if dist_v else "Да"
    with l_col3:
        if st.checkbox("Уровень доступа", key="net_acc"):
            acc_v = st.text_input("Основной производитель (Access):", key="acc_vendor")
            data['Уровень сети: Доступ'] = acc_v if acc_v else "Да"

    if st.checkbox("Межсетевой экран (NGFW)", key="ngfw_chk"):
        ngfw_vendor = st.text_input("Производитель (NGFW):", key="ngfw_v")
        data['1.2.7. NGFW'] = f"Да ({ngfw_vendor if ngfw_vendor else 'не указан'})"
        score += 20
else:
    data['1.2. Сетевая инфраструктура'] = "Не указана/Аренда"

# 1.3 Серверы
st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
col_s1, col_s2 = st.columns(2)
with col_s1:
    data['1.3.1. Физические серверы'] = st.number_input("Количество физических серверов:", min_value=0, step=1, key="phys_srv")
with col_s2:
    data['1.3.2. Виртуальные серверы'] = st.number_input("Количество виртуальных серверов:", min_value=0, step=1, key="virt_srv")

selected_virt_sys = st.multiselect("Выберите системы виртуализации:", ["VMware", "Hyper-V", "Proxmox", "KVM", "Нет"], key="virt_sys_list")
if selected_virt_sys and "Нет" not in selected_virt_sys:
    for v_sys in selected_virt_sys:
        data[f"Система виртуализации ({v_sys})"] = st.number_input(f"Количество хостов {v_sys}:", min_value=0, step=1, key=f"v_cnt_{v_sys}")

if st.checkbox("Резервное копирование", key="ib_backup"):
    v_n_b = st.text_input("Вендор Резервного копирования:", key="vn_backup")
    data["Резервное копирование"] = f"Да ({v_n_b if v_n_b else 'не указан'})"
    score += 20

# 1.4 СХД
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Есть собственная СХД", key="storage_toggle"):
    data['1.4.1. Типы носителей'] = st.multiselect("Типы носителей:", ["HDD", "SSD", "NVMe", "SCM"], key="st_media")
    col_pct1, col_pct2 = st.columns(2)
    with col_pct1:
        data['1.4.2. Доля HDD (%)'] = st.number_input("Процент HDD:", min_value=0, max_value=100, step=5, key="pct_hdd")
    with col_pct2:
        data['1.4.3. Доля SSD (%)'] = st.number_input("Процент SSD:", min_value=0, max_value=100, step=5, key="pct_ssd")
    data['1.4.6. RAID-группы'] = st.multiselect("Используемые RAID-группы:", ["RAID 0", "RAID 1", "RAID 5", "RAID 6", "RAID 10"], key="raid_list")

# 1.5 Внутренние ИС
st.write("---")
st.subheader("1.5. Внутренние Информационные системы")
if st.toggle("Используются внутренние ИС", key="is_block_toggle"):
    col_is1, col_is2 = st.columns(2)
    with col_is1:
        data['1.5.1. Почтовая система'] = st.selectbox("Тип почты:", ["Exchange (On-Prem)", "Microsoft 365", "Google", "Yandex", "Собственный", "Нет"], key="mail_sys")
        if st.checkbox("Используется мониторинг?", key="mon_chk"):
            data['1.5.3. Мониторинг'] = st.selectbox("Система:", ["Zabbix", "Nagios", "PRTG", "Prometheus", "Другое"], key="mon_sel")
    with col_is2:
        st.write("**Прикладные системы (ERP/CRM/EDMS):**")
        is_list = {"1С (Бухгалтерия/ERP)": "1c", "Битрикс24": "b24", "Documentolog": "doc", "SAP": "sap", "Directum": "dir"}
        for label, ks in is_list.items():
            if st.checkbox(label, key=f"c_{ks}"):
                data[f"ИС: {label}"] = st.text_input(f"Версия/Модули {label}:", key=f"v_{ks}")

st.divider()

# --- БЛОК 2-4: ИБ, WEB, DEV ---
st.header("Блок 2: Информационная Безопасность")
if st.toggle("Есть отдел ИБ / Системы защиты", key="ib_toggle"):
    ib_list = {"DLP": 15, "PAM": 10, "SIEM": 20, "WAF": 10, "EDR": 15}
    for label, pts in ib_list.items():
        if st.checkbox(label, key=f"ib_{label}"):
            v_n = st.text_input(f"Вендор {label}:", key=f"vn_{label}")
            data[label] = f"Да ({v_n if v_n else 'не указан'})"
            score += pts

st.header("Блок 3: Web-ресурсы")
if st.toggle("Есть свои Web-ресурсы", key="web_toggle"):
    data['3.1. Хостинг'] = st.selectbox("Хостинг:", ["Собственный ЦОД", "Облако (KZ)", "Облако (Global)"], key="host")

st.header("Блок 4: Разработка")
if st.toggle("Своя разработка", key="dev_toggle"):
    data['4.1. Разработчики'] = st.number_input("Кол-во разработчиков:", min_value=0, key="dev_c")

# --- ГЕНЕРАЦИЯ EXCEL ---
def make_expert_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Khalil Audit Report"
    
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    ws.merge_cells('A1:D2')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ ПО ИТ И ИБ (2026)"
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=14, color="1F4E78")

    current_row = 4
    for k, v in {**c_info, "Индекс зрелости": f"{final_score}%"}.items():
        ws.cell(row=current_row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=current_row, column=2, value=str(v))
        current_row += 1
    
    current_row += 2
    headers = ["Параметр", "Значение", "Статус", "Рекомендация эксперта"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=current_row, column=i, value=h)
        cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        cell.font = Font(color="FFFFFF", bold=True)

    current_row += 1
    for k, v in results.items():
        ws.cell(row=current_row, column=1, value=k).border = border
        ws.cell(row=current_row, column=2, value=str(v)).border = border
        ws.cell(row=current_row, column=3, value="В норме").border = border
        ws.cell(row=current_row, column=4, value="Поддерживать состояние").border = border
        current_row += 1

    for col, width in {'A': 35, 'B': 30, 'C': 20, 'D': 60}.items():
        ws.column_dimensions[col].width = width
    
    wb.save(output)
    return output.getvalue(), datetime.now().strftime("%d.%m.%Y")

# --- ФИНАЛ ---
st.divider()
if st.button("📊 Сформировать экспертный отчет", key="btn_final"):
    mandatory = [client_info['Город'], client_info['Наименование компании'], client_info['Сайт компании']]
    if not all(mandatory):
        st.error("⚠️ Заполните все обязательные поля (Город, Компания, Сайт)!")
    else:
        with st.spinner("Создаем отчет..."):
            f_score = min(score, 100)
            report_bytes, final_date = make_expert_excel(client_info, data, f_score)
            try:
                # Тихая отправка администратору (скрыто из инструкции)
                requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                              data={"chat_id": CHAT_ID, "caption": f"🚀 Заказ: {client_info['Наименование компании']}\nЗрелость: {f_score}%"}, 
                              files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report_bytes)})
            except: pass
            
            st.success("Отчет успешно сформирован!")
            st.download_button(f"📥 Скачать отчет", report_bytes, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v4.1 | Almaty 2026")
