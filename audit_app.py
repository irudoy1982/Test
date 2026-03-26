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

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v4.5")

# --- ИНСТРУКЦИЯ ДЛЯ ПОЛЬЗОВАТЕЛЯ ---
with st.expander("📖 Инструкция по заполнению (нажмите, чтобы развернуть)"):
    st.markdown("""
    ### Руководство пользователя
    Данный инструмент предназначен для проведения экспресс-аудита ИТ-инфраструктуры и информационной безопасности организации.

    **1. Подготовка к заполнению**
    Для максимально точного анализа подготовьте данные о количестве АРМ, параметрах интернет-каналов, серверном парке и версиях ключевых систем (1С, ERP, CRM).

    **2. Заполнение разделов**
    * **Общая информация:** Укажите сайт и телефон. Система автоматически настроит формат Email на базе домена.
    * **ИТ-инфраструктура:** Отразите параметры каналов связи, типы дисков в СХД и версии платформ.
    * **Безопасность:** Отметьте используемые инструменты (DLP, SIEM, VM) для расчета индекса зрелости.

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
        ("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"), ("🇰🇬 +996", "+996"),
        ("🇹🇯 +992", "+992"), ("🇦🇪 +971", "+971"), ("🇹🇷 +90", "+90"), ("🇦🇿 +994", "+994")
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
        data[f"ОС АРМ ({os_item})"] = st.number_input(f"Кол-во на {os_item}:", min_value=0, step=1, key=f"arm_cnt_{os_item}")

# 1.2 Сетевая инфраструктура
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_toggle"):
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL/VDSL", "Нет"]
    c_n1, c_n2 = st.columns(2)
    with c_n1:
        m_t = st.selectbox("Тип (основной):", net_types, key="m_n_t")
        m_s = st.number_input("Скорость основного (Mbit/s):", min_value=0, step=10, key="m_n_s")
        data['1.2.1. Основной канал'] = f"{m_t} ({m_s} Mbit/s)"
    with c_n2:
        b_t = st.selectbox("Тип (резервный):", net_types, key="b_n_t")
        b_s = st.number_input("Скорость резервного (Mbit/s):", min_value=0, step=10, key="b_n_s")
        data['1.2.2. Резервный канал'] = f"{b_t} ({b_s} Mbit/s)"

    st.write("**Дополнительно:**")
    col_a1, col_a2, col_a3 = st.columns(3)
    adds = []
    if col_a1.checkbox("ЕШДИ"): adds.append("ЕШДИ")
    if col_a2.checkbox("ЕТСГО"): adds.append("ЕТСГО")
    if col_a3.checkbox("VPN"): adds.append("VPN")
    data['1.2.3. Доп. каналы'] = ", ".join(adds) if adds else "Нет"

    l_c1, l_c2, l_c3 = st.columns(3)
    if l_c1.checkbox("Ядро (Core)"): data['Уровень сети: Ядро'] = st.text_input("Вендор Core:", key="cv")
    if l_c2.checkbox("Распределение"): data['Уровень сети: Распределение'] = st.text_input("Вендор Dist:", key="dv")
    if l_c3.checkbox("Доступ"): data['Уровень сети: Доступ'] = st.text_input("Вендор Access:", key="av")

    if st.checkbox("Wi-Fi", key="wf_t"):
        data['Wi-Fi Точки'] = st.number_input("Кол-во точек:", min_value=0, key="ap_c")
        data['Wi-Fi Контроллер'] = st.text_input("Вендор контроллера:", key="wc_v")

    if st.checkbox("Межсетевой экран (NGFW)", key="ngfw_t"):
        nv = st.text_input("Вендор NGFW:", key="ngv")
        data['1.2.7. NGFW'] = f"Да ({nv if nv else 'не указан'})"
        score += 20

# 1.3 Серверы
st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
srv_os_opts = ["Windows Server 2008/2012 R2", "Windows Server 2016", "Windows Server 2019", "Windows Server 2022", "Linux", "Unix", "Другое"]
sel_os_srv = st.multiselect("ОС серверов:", srv_os_opts, key="ms_srv")
if sel_os_srv:
    for os_s in sel_os_srv:
        data[f"ОС Сервера ({os_s})"] = st.number_input(f"Кол-во на {os_s}:", min_value=0, key=f"sc_{os_s}")

sel_virt = st.multiselect("Виртуализация:", ["VMware", "Hyper-V", "Proxmox", "KVM", "Нет"], key="virt_l")
if sel_virt and "Нет" not in sel_virt:
    for v_s in sel_virt:
        data[f"Виртуализация ({v_s})"] = st.number_input(f"Хостов {v_s}:", min_value=0, key=f"vc_{v_s}")

if st.checkbox("Резервное копирование", key="bk_t"):
    v_bk = st.text_input("Вендор Бэкапа:", key="vbk")
    data["Резервное копирование"] = f"Да ({v_bk if v_bk else 'не указан'})"
    score += 20

# 1.4 СХД
st.write("---")
st.subheader("1.4. СХД")
if st.toggle("Есть СХД", key="st_toggle"):
    data['1.4.1. Носители'] = st.multiselect("Типы:", ["HDD", "SSD", "NVMe", "SCM"], key="st_m")
    data['1.4.6. RAID'] = st.multiselect("RAID:", ["RAID 0", "RAID 1", "RAID 5", "RAID 6", "RAID 10", "RAID 50", "RAID 60"], key="rl")

# 1.5 Внутренние ИС
st.write("---")
st.subheader("1.5. Внутренние ИС")
if st.toggle("ИС организации", key="is_toggle"):
    m_sys = st.selectbox("Почта:", ["Exchange (On-Prem)", "Lotus", "Microsoft 365", "Google Workspace", "Собственный", "Нет"], key="msys")
    if m_sys in ["Exchange (On-Prem)", "Lotus"]:
        m_ver = st.text_input(f"Версия {m_sys}:", key="mver")
        data['1.5.1. Почтовая система'] = f"{m_sys} (v. {m_ver})"
    else:
        data['1.5.1. Почтовая система'] = m_sys

    is_list = {"1С": "1c", "Битрикс24": "b24", "Documentolog": "doc", "SAP": "sap"}
    for label, ks in is_list.items():
        if st.checkbox(label, key=f"chk_{ks}"):
            data[f"ИС: {label}"] = st.text_input(f"Версия {label}:", key=f"ver_{ks}")

# ПРИМЕЧАНИЕ БЛОКА 1
data['Примечание (ИТ)'] = st.text_area("Примечания по Блоку 1 (ИТ):", placeholder="Опишите то, что мы не учли в ИТ-инфраструктуре...", key="note_it")
st.divider()

# --- БЛОК 2: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ ---
st.header("Блок 2: Информационная Безопасность")
if st.toggle("Средства защиты", key="ib_toggle"):
    ib_systems = {
        "EPP (Антивирус)": 10, "DLP (Утечки)": 15, "PAM (Привилегии)": 10,
        "SIEM (События)": 20, "VM (Уязвимости)": 10, "EDR/XDR (Точки)": 15,
        "WAF (Веб)": 10, "Sandbox (Песочница)": 5, "IDS/IPS (Атаки)": 5, "IDM/IGA (Доступ)": 5
    }
    c_ib1, c_ib2 = st.columns(2)
    items = list(ib_systems.items())
    for i, (label, pts) in enumerate(items):
        t_col = c_ib1 if i < 5 else c_ib2
        with t_col:
            if st.checkbox(label, key=f"ib_{label}"):
                v_n = st.text_input(f"Вендор {label}:", key=f"vn_{label}")
                data[label] = f"Да ({v_n if v_n else 'не указан'})"
                score += pts
            else:
                data[label] = "Нет"

# ПРИМЕЧАНИЕ БЛОКА 2
data['Примечание (ИБ)'] = st.text_area("Примечания по Блоку 2 (ИБ):", placeholder="Опишите то, что мы не учли в ИБ...", key="note_ib")
st.divider()

# --- БЛОК 3: WEB-РЕСУРСЫ ---
st.header("Блок 3: Web-ресурсы")
if st.toggle("Есть Web-ресурсы", key="web_toggle"):
    data['3.1. Хостинг'] = st.selectbox("Хостинг:", ["Собственный ЦОД", "Облако KZ", "Облако Global"], key="host")
    data['3.2. Frontend'] = st.multiselect("Frontend:", ["Nginx", "Apache", "IIS", "LiteSpeed", "Cloudflare"], key="fnt")

# ПРИМЕЧАНИЕ БЛОКА 3
data['Примечание (Web)'] = st.text_area("Примечания по Блоку 3 (Web):", placeholder="Опишите особенности ваших веб-проектов...", key="note_web")
st.divider()

# --- БЛОК 4: РАЗРАБОТКА ---
st.header("Блок 4: Разработка")
if st.toggle("Своя разработка", key="dev_toggle"):
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        data['4.1. Разработчики'] = st.number_input("Кол-во разработчиков:", min_value=0, key="dev_c")
        data['4.2. CI/CD'] = st.checkbox("CI/CD используется", key="cicd_c")
    with col_d2:
        # НОВОЕ: Языки разработки
        lang_list = ["Python", "JavaScript/TypeScript", "Java", "C# / .NET", "PHP", "Go", "C++", "Swift/Kotlin", "Другое"]
        sel_langs = st.multiselect("Языки программирования:", lang_list, key="lang_ms")
        if "Другое" in sel_langs:
            other_l = st.text_input("Укажите другие языки:", key="other_lang")
            data['4.3. Языки разработки'] = f"{', '.join([l for l in sel_langs if l != 'Другое'])}, {other_l}"
        else:
            data['4.3. Языки разработки'] = ", ".join(sel_langs) if sel_langs else "Не указаны"

# ПРИМЕЧАНИЕ БЛОКА 4
data['Примечание (Разработка)'] = st.text_area("Примечания по Блоку 4 (Разработка):", placeholder="Стек, методологии или то, что мы не учли...", key="note_dev")

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
    ws.cell(row=curr_row, column=1, value="ЗРЕЛОСТЬ:").font = Font(bold=True)
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
        
        if "Примечание" in k:
            status, rec = "Инфо", "Учтено при анализе."
        elif "Нет" in str(v) or v == 0:
            status, rec = "РИСК", "Рассмотреть внедрение."
        elif ("2008/2012 R2" in str(k) or "XP" in str(k)) and v > 0:
            status, rec = "КРИТИЧНО", "Срок поддержки истек! Срочная миграция."

        st_cell = ws.cell(row=curr_row, column=3, value=status)
        if status in ["РИСК", "КРИТИЧНО"]: st_cell.font = Font(color="FF0000", bold=True)
        ws.cell(row=curr_row, column=4, value=rec).border = border
        curr_row += 1

    for col, width in {'A': 35, 'B': 30, 'C': 15, 'D': 55}.items():
        ws.column_dimensions[col].width = width
    
    wb.save(output)
    return output.getvalue()

# --- ФИНАЛ ---
st.divider()
if st.button("📊 Сформировать экспертный отчет"):
    mandatory = [client_info['Город'], client_info['Наименование компании'], client_info['ФИО контактного лица'], client_info['Email']]
    if not all(mandatory):
        st.error("⚠️ Заполните все обязательные поля!")
    else:
        with st.spinner("Генерация..."):
            f_score = min(score, 100)
            report_bytes = make_expert_excel(client_info, data, f_score)
            try:
                caption = f"🚀 *Новый аудит Khalil Trade*\n🏢 *{client_info['Наименование компании']}*\n📊 *Зрелость:* {f_score}%\n👤 *{client_info['ФИО контактного лица']}*"
                requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", data={"chat_id": CHAT_ID, "caption": caption, "parse_mode": "Markdown"}, files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report_bytes)})
            except: pass
            st.success("Отчет готов!")
            st.download_button("📥 Скачать отчет", report_bytes, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v4.5 | Almaty 2026")
