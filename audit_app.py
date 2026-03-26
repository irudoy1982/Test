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

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v4.9")

# --- ИНСТРУКЦИЯ ДЛЯ ПОЛЬЗОВАТЕЛЯ ---
with st.expander("📖 Инструкция по заполнению (нажмите, чтобы развернуть)"):
    st.markdown("""
    ### Руководство пользователя
    Данный инструмент предназначен для проведения экспресс-аудита ИТ-инфраструктуры и информационной безопасности организации.

    **1. Подготовка к заполнению**
    Подготовьте данные о количестве АРМ, параметрах интернет-каналов, серверном парке и сетевом оборудовании.

    **2. Заполнение разделов**
    * **Общая информация:** Укажите сайт и телефон. Система настроит Email на базе домена.
    * **ИТ-инфраструктура:** Отразите параметры связи, оборудование и версии ОС.
    * **Безопасность:** Отметьте инструменты защиты для расчета индекса зрелости.
    * **Примечания:** В конце активных блоков можно добавить текстовые пояснения.

    **3. Получение результата**
    * Нажмите **«Сформировать экспертный отчет»**, затем **«📥 Скачать отчет»**. 
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
    country_codes = [
        ("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"), ("🇰🇬 +996", "+996"),
        ("🇹🇯 +992", "+992"), ("🇦🇪 +971", "+971"), ("🇹🇷 +90", "+90"), ("🇦🇿 +994", "+994")
    ]
    selected_code = p_col1.selectbox("Код", country_codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed")
    client_info['Контактный телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

st.divider()

# --- БЛОК 1: ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ ---
st.header("Блок 1: Информационные технологии")

# 1.1 Конечные точки
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)
data['1.1. Всего АРМ'] = total_arm
selected_os_arm = st.multiselect("ОС на АРМ:", ["Windows XP/Vista/7/8", "Windows 10", "Windows 11", "Linux", "macOS", "Другое"])
if selected_os_arm:
    for os_item in selected_os_arm:
        data[f"ОС АРМ ({os_item})"] = st.number_input(f"Кол-во на {os_item}:", min_value=0, step=1, key=f"arm_{os_item}")

# 1.2 Сетевая инфраструктура (ВЗЯТО ИЗ V4.2)
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

    st.write("**Дополнительные каналы:**")
    col_add1, col_add2, col_add3 = st.columns(3)
    add_channels = []
    if col_add1.checkbox("ЕШДИ", key="chk_eshdi"): add_channels.append("ЕШДИ")
    if col_add2.checkbox("ЕТСГО", key="chk_etsgo"): add_channels.append("ЕТСГО")
    if col_add3.checkbox("VPN", key="chk_vpn"): add_channels.append("VPN")
    data['1.2.3. Доп. каналы'] = ", ".join(add_channels) if add_channels else "Нет"

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

    st.write("**Беспроводная сеть:**")
    if st.checkbox("Wi-Fi", key="wifi_toggle"):
        w_col1, w_col2, w_col3 = st.columns(3)
        with w_col1:
            if st.checkbox("Контроллер", key="wifi_ctrl"):
                wc_v = st.text_input("Производитель/модель контроллера:", key="wc_vendor")
                data['Wi-Fi Контроллер'] = wc_v if wc_v else "Да"
        with w_col2:
            data['Wi-Fi Точки доступа'] = st.number_input("Количество точек доступа (шт):", min_value=0, step=1, key="ap_cnt")
        with w_col3:
            wf_types = ["Wi-Fi 6/6E (802.11ax)", "Wi-Fi 5 (802.11ac)", "Wi-Fi 4 (802.11n)", "Другое"]
            data['Wi-Fi Тип'] = st.selectbox("Тип Wi-Fi:", wf_types, key="wf_type_sel")

    if st.checkbox("Межсетевой экран (NGFW)", key="ngfw_chk"):
        ngfw_vendor = st.text_input("Производитель (NGFW):", key="ngfw_v")
        data['1.2.7. NGFW'] = f"Да ({ngfw_vendor if ngfw_vendor else 'не указан'})"
        score += 20

# 1.3 Серверы
st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
s_os_list = ["Windows Server 2008/2012 R2", "Windows Server 2016", "Windows Server 2019", "Windows Server 2022", "Linux", "Unix", "Другое"]
selected_os_srv = st.multiselect("ОС серверов:", s_os_list)
if selected_os_srv:
    for os_s in selected_os_srv:
        data[f"ОС Сервера ({os_s})"] = st.number_input(f"Кол-во на {os_s}:", min_value=0, key=f"srv_{os_s}")

selected_virt = st.multiselect("Виртуализация:", ["VMware", "Hyper-V", "Proxmox", "KVM", "Нет"])
if selected_virt and "Нет" not in selected_virt:
    for v_s in selected_virt:
        data[f"Виртуализация ({v_s})"] = st.number_input(f"Хостов {v_s}:", min_value=0, key=f"virt_{v_s}")

if st.checkbox("Резервное копирование"):
    data["Бэкап система"] = st.text_input("Вендор Бэкапа (Veeam и др.):")
    score += 20

# 1.4 СХД (ВЗЯТО ИЗ V4.2)
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
    data['1.4.4. Гибридная СХД'] = col_chk1.checkbox("Используется гибридная СХД", key="hybrid_st")
    data['1.4.5. All-Flash'] = col_chk2.checkbox("Есть All-Flash массивы", key="allflash_st")
    data['1.4.6. RAID-группы'] = st.multiselect("Используемые RAID-группы:", ["RAID 0", "RAID 1", "RAID 5", "RAID 6", "RAID 10", "RAID 50", "RAID 60", "JBOD"], key="raid_list")

# 1.5 Внутренние ИС
st.write("---")
st.subheader("1.5. Внутренние Информационные системы")
if st.toggle("ИС организации", key="is_toggle"):
    m_opts = ["Exchange (On-Prem)", "Lotus", "Microsoft 365", "Google Workspace", "Собственный", "Нет"]
    m_sys = st.selectbox("Почта:", m_opts)
    if m_sys in ["Exchange (On-Prem)", "Lotus"]:
        m_ver = st.text_input(f"Версия {m_sys}:")
        data['1.5.1. Почтовая система'] = f"{m_sys} (v. {m_ver})"
    else:
        data['1.5.1. Почтовая система'] = m_sys

    is_list = {"1С": "1c", "Битрикс24": "b24", "Documentolog": "doc", "SAP": "sap"}
    for label, ks in is_list.items():
        if st.checkbox(label): data[f"ИС: {label}"] = st.text_input(f"Версия {label}:")

    data['Примечание (ИТ)'] = st.text_area("Примечания по ИТ-инфраструктуре:", placeholder="Доп. информация...")

st.divider()

# --- БЛОК 2: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ ---
st.header("Блок 2: Информационная Безопасность")
if st.toggle("Средства защиты", key="ib_toggle"):
    ib_systems = {
        "EPP (Антивирус)": 10, "DLP (Утечки)": 15, "PAM (Привилегии)": 10,
        "SIEM (Мониторинг)": 20, "VM (Уязвимости)": 10, "EDR/XDR (Точки)": 15,
        "WAF (Веб)": 10, "Sandbox (Песочница)": 5, "IDS/IPS (Атаки)": 5, "IDM/IGA (Доступ)": 5
    }
    col_ib1, col_ib2 = st.columns(2)
    items = list(ib_systems.items())
    for i, (label, pts) in enumerate(items):
        target_col = col_ib1 if i < 5 else col_ib2
        with target_col:
            if st.checkbox(label, key=f"ib_{label}"):
                v_n = st.text_input(f"Вендор {label}:", key=f"vn_{label}")
                data[label] = f"Да ({v_n if v_n else 'не указан'})"
                score += pts
            else:
                data[label] = "Нет"
    data['Примечание (ИБ)'] = st.text_area("Примечания по ИБ:", placeholder="Особенности защиты...")

st.divider()

# --- БЛОК 3: WEB-РЕСУРСЫ ---
st.header("Блок 3: Web-ресурсы")
if st.toggle("Web-ресурсы", key="web_toggle"):
    data['3.1. Хостинг'] = st.selectbox("Хостинг:", ["Собственный ЦОД", "Облако KZ", "Облако Global"])
    data['3.2. Frontend'] = st.multiselect("Frontend серверы:", ["Nginx", "Apache", "IIS", "LiteSpeed", "Cloudflare"])
    data['Примечание (Web)'] = st.text_area("Примечания по Web:", placeholder="CMS и т.д.")

st.divider()

# --- БЛОК 4: РАЗРАБОТКА ---
st.header("Блок 4: Разработка")
if st.toggle("Разработка", key="dev_toggle"):
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        data['4.1. Разработчики'] = st.number_input("Кол-во разработчиков:", min_value=0)
        data['4.2. CI/CD'] = st.checkbox("Используется CI/CD")
    with col_d2:
        lang_list = ["Python", "JavaScript/TypeScript", "Java", "C# / .NET", "PHP", "Go", "C++", "Swift/Kotlin", "Другое"]
        sel_langs = st.multiselect("Языки программирования:", lang_list)
        if "Другое" in sel_langs:
            other_l = st.text_input("Укажите другие языки:")
            data['4.3. Языки разработки'] = f"{', '.join([l for l in sel_langs if l != 'Другое'])}, {other_l}"
        else:
            data['4.3. Языки разработки'] = ", ".join(sel_langs) if sel_langs else "Не указаны"
    data['Примечание (Разработка)'] = st.text_area("Примечания по Разработке:", placeholder="Стек, методологии...")

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

    curr_row = 4
    for k, v in c_info.items():
        ws.cell(row=curr_row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=curr_row, column=2, value=str(v))
        curr_row += 1
    
    auto_date = datetime.now().strftime("%d.%m.%Y %H:%M")
    ws.cell(row=curr_row, column=1, value="ИНДЕКС ЗРЕЛОСТИ:").font = Font(bold=True)
    score_cell = ws.cell(row=curr_row, column=2, value=f"{final_score}%")
    bg = "92D050" if final_score > 70 else "FFC000" if final_score > 40 else "FF7C80"
    score_cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
    
    curr_row += 3
    headers = ["Параметр", "Значение", "Статус", "Рекомендация"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=curr_row, column=i, value=h)
        cell.fill = header_fill; cell.font = white_font

    curr_row += 1
    for k, v in results.items():
        ws.cell(row=curr_row, column=1, value=k).border = border
        ws.cell(row=curr_row, column=2, value=str(v)).border = border
        status, rec = "В норме", "Поддерживать состояние."
        is_risk = False
        if "Примечание" in k:
            status, rec = "Инфо", "Учтено при анализе."
        elif "Нет" in str(v) or v == 0:
            is_risk = True; status = "РИСК"; rec = "Рассмотреть внедрение."
        if ("2008/2012 R2" in str(k) or "XP" in str(k)) and v > 0:
            is_risk = True; status = "КРИТИЧНО"; rec = "Срок поддержки истек. Срочная миграция!"
            
        st_cell = ws.cell(row=curr_row, column=3, value=status)
        if is_risk: st_cell.font = Font(color="FF0000", bold=True)
        ws.cell(row=curr_row, column=4, value=rec).border = border
        curr_row += 1

    for col, width in {'A': 35, 'B': 30, 'C': 15, 'D': 55}.items():
        ws.column_dimensions[col].width = width
    
    wb.save(output)
    return output.getvalue(), auto_date

# --- ФИНАЛ ---
st.divider()
if st.button("📊 Сформировать экспертный отчет"):
    mandatory = [client_info['Город'], client_info['Наименование компании'], client_info['ФИО контактного лица'], client_info['Email']]
    if not all(mandatory):
        st.error("⚠️ Заполните все обязательные поля!")
    else:
        with st.spinner("Обработка..."):
            f_score = min(score, 100)
            report_bytes, final_date = make_expert_excel(client_info, data, f_score)
            try:
                caption = f"🚀 *Новый аудит Khalil Trade*\n🏢 *{client_info['Наименование компании']}*\n📊 *Зрелость:* {f_score}%\n👤 *{client_info['ФИО контактного лица']}*"
                requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", data={"chat_id": CHAT_ID, "caption": caption, "parse_mode": "Markdown"}, files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report_bytes)})
            except: pass
            st.success("Отчет готов!")
            st.download_button("📥 Скачать отчет", report_bytes, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v4.9 | Almaty 2026")
