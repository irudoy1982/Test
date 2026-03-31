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

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v6.7")

# --- ИНСТРУКЦИЯ ---
with st.expander("📖 Инструкция по заполнению"):
    st.markdown("""
    1. **Общая информация:** Поля со звездочкой (*) обязательны.
    2. **Логический контроль:** Система проверяет соответствие ОС количеству устройств.
    3. **Результат:** Вы получите Excel-файл с анализом рисков.
    """)

data = {}
client_info = {}
validation_errors = []
score = 0

# --- ШАПКА: ИНФОРМАЦИЯ О КЛИЕНТЕ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город*")
    client_info['Наименование компании'] = st.text_input("Наименование компании*")
    site_input = st.text_input("Сайт компании*", key="site_field", placeholder="example.kz")
    client_info['Сайт компании'] = site_input

    custom_email_mode = st.checkbox("Email отличается от сайта")
    clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
    
    if custom_email_mode:
        client_info['Email'] = st.text_input("Email контактного лица*")
    elif clean_domain and "." in clean_domain:
        st.write("Email контактного лица (логин)*")
        e_col1, e_col2 = st.columns([2, 3])
        with e_col1:
            email_prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="email_pre")
        with e_col2:
            st.markdown(f"<div style='padding-top: 5px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
        client_info['Email'] = f"{email_prefix}@{clean_domain}" if email_prefix else ""
    else:
        client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица*")
    client_info['Должность'] = st.text_input("Должность*")
    st.write("Контактный телефон*")
    p_col1, p_col2 = st.columns([1, 2])
    country_codes = [("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"), ("🇦🇪 +971", "+971")]
    selected_code = p_col1.selectbox("Код", country_codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed")
    client_info['Контактный телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

# Валидация
mandatory = ['Город', 'Наименование компании', 'Сайт компании', 'Email', 'ФИО контактного лица', 'Должность']
for f in mandatory:
    if not client_info.get(f): validation_errors.append(f"Не заполнено: {f}")

st.divider()

# --- БЛОК 1: ИТ ---
st.header("Блок 1: Информационные технологии")

# 1.1 АРМ
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт)", min_value=0, step=1)
data['1.1. Всего АРМ'] = total_arm
os_arm_list = ["Windows XP/7/8", "Windows 10", "Windows 11", "Linux", "macOS", "Другое"]
selected_os_arm = st.multiselect("ОС на АРМ", os_arm_list)
sum_os_arm = 0
if selected_os_arm:
    for os_i in selected_os_arm:
        v = st.number_input(f"Кол-во на {os_i}", min_value=0, key=f"arm_{os_i}")
        data[f"ОС АРМ ({os_i})"] = v
        sum_os_arm += v
if total_arm > 0 and sum_os_arm != total_arm:
    st.warning(f"⚠️ Ошибка: Всего {total_arm}, по ОС {sum_os_arm}")
    validation_errors.append("Несовпадение АРМ и ОС")

# 1.2 Сеть (ВОССТАНОВЛЕНО ПОЛНОСТЬЮ)
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_tgl"):
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL/VDSL", "Нет"]
    c1, c2 = st.columns(2)
    with c1:
        m_t = st.selectbox("Тип (основной)", net_types)
        m_s = st.number_input("Скорость основного (Mbps)", min_value=0)
        data['1.2.1. Основной канал'] = f"{m_t} ({m_s} Mbps)"
    with c2:
        b_t = st.selectbox("Тип (резервный)", net_types, index=6)
        b_s = st.number_input("Скорость резервного (Mbps)", min_value=0)
        data['1.2.2. Резервный канал'] = b_t if b_t == "Нет" else f"{b_t} ({b_s} Mbps)"

    st.write("**Дополнительные каналы**")
    ca1, ca2, ca3 = st.columns(3)
    adds = []
    if ca1.checkbox("ЕШДИ"): adds.append("ЕШДИ")
    if ca2.checkbox("ЕТСГО"): adds.append("ЕТСГО")
    if ca3.checkbox("VPN"): adds.append("VPN")
    data['1.2.3. Доп. каналы'] = ", ".join(adds) if adds else "Нет"

    st.write("**Уровни сети и оборудование**")
    cl1, cl2, cl3 = st.columns(3)
    with cl1:
        if st.checkbox("Ядро (Core)"):
            cv = st.text_input("Вендор Core")
            data['Уровень Core'] = cv
    with cl2:
        if st.checkbox("Распределение"):
            dv = st.text_input("Вендор Dist")
            data['Уровень Dist'] = dv
    with cl3:
        if st.checkbox("Доступ (Access)"):
            av = st.text_input("Вендор Access")
            data['Уровень Access'] = av

    if st.checkbox("Wi-Fi"):
        w1, w2, w3 = st.columns(3)
        with w1:
            wf_v = st.text_input("Вендор/Контроллер Wi-Fi")
        with w2:
            wf_q = st.number_input("Кол-во точек", min_value=0)
        with w3:
            wf_st = st.selectbox("Стандарт", ["Wi-Fi 6/6E", "Wi-Fi 5", "Wi-Fi 4"])
        data['1.2.4. Wi-Fi'] = f"{wf_v}, {wf_q} шт, {wf_st}"

# 1.3 Серверы (ВОССТАНОВЛЕНО ПОЛНОСТЬЮ)
st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
cs1, cs2 = st.columns(2)
with cs1: phys = st.number_input("Физические серверы", min_value=0)
with cs2: virt = st.number_input("Виртуальные серверы", min_value=0)
data['1.3.1. Физика'] = phys; data['1.3.2. Виртуализация'] = virt

s_os = ["Win Serv 2008/2012 R2", "Win Serv 2016/2019", "Win Serv 2022", "Linux", "Unix"]
sel_srv_os = st.multiselect("ОС Серверов", s_os)
sum_srv_os = 0
for os_s in sel_srv_os:
    v_os = st.number_input(f"Кол-во {os_s}", min_value=0, key=f"sos_{os_s}")
    sum_srv_os += v_os

sel_virt = st.multiselect("Системы виртуализации", ["VMware", "Hyper-V", "Proxmox", "KVM", "Другое"])
for vsys in sel_virt:
    vh = st.number_input(f"Хостов {vsys}", min_value=0, key=f"vsys_{vsys}")
    data[f"Виртуализация ({vsys})"] = vh

# 1.4 СХД (ВОССТАНОВЛЕНО ПОЛНОСТЬЮ)
st.write("---")
st.subheader("1.4. СХД")
if st.toggle("Есть СХД"):
    st_v = st.text_input("Вендор СХД")
    st_c = st.number_input("Объем (TB)", min_value=0)
    st_m = st.multiselect("Носители", ["HDD", "SSD", "NVMe"])
    st_r = st.multiselect("RAID", ["RAID 5", "RAID 6", "RAID 10", "RAID 60"])
    data['1.4. СХД'] = f"{st_v}, {st_c} TB, {', '.join(st_m)}, RAID: {', '.join(st_r)}"

# 1.5 ИС (ВОССТАНОВЛЕНО ПОЛНОСТЬЮ)
st.write("---")
st.subheader("1.5. Внутренние ИС")
if st.toggle("ИС Организации"):
    mail = st.selectbox("Почта", ["Exchange On-Prem", "M365", "Google", "Собственный"])
    if "On-Prem" in mail:
        m_v = st.text_input("Версия почты")
        data['Почта'] = f"{mail} ({m_v})"
    else: data['Почта'] = mail
    
    for app in ["1С", "Битрикс24", "Documentolog", "SAP"]:
        if st.checkbox(app):
            av = st.text_input(f"Версия {app}")
            data[f"ИС {app}"] = av

# --- БЛОК 2: ИБ ---
st.write("---")
st.header("Блок 2: Информационная Безопасность")
ib_sys = {"EPP": 10, "DLP": 15, "PAM": 10, "SIEM": 20, "MFA": 15, "WAF": 10, "Anti-DDoS": 15}
col_ib1, col_ib2 = st.columns(2)
for i, (name, pts) in enumerate(ib_sys.items()):
    with (col_ib1 if i < 4 else col_ib2):
        if st.checkbox(name):
            v_ib = st.text_input(f"Вендор {name}", key=f"vib_{name}")
            data[name] = f"Да ({v_ib})"
            score += pts
        else: data[name] = "Нет"

# --- БЛОК 4: РАЗРАБОТКА (ВОССТАНОВЛЕНО ПОЛНОСТЬЮ) ---
st.write("---")
st.header("Блок 4: Разработка")
if st.toggle("Разработка"):
    d_q = st.number_input("Кол-во разработчиков", min_value=0)
    d_langs = st.multiselect("Языки", ["Python", "JS/TS", "Java", "C#", "PHP", "Go", "Другое"])
    if d_q > 0 and not d_langs: st.info("Укажите языки.")
    if d_q == 0 and d_langs: validation_errors.append("Противоречие в разработке")
    data['Разработка'] = f"{d_q} чел, {', '.join(d_langs)}"

# --- EXCEL ---
def make_report(c_info, results, f_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Report"
    
    # Заголовок и цвета (как в исходнике)
    ws.merge_cells('A1:D1')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ 2026"
    ws['A1'].font = Font(bold=True, size=14)
    
    row = 3
    for k, v in c_info.items():
        ws.cell(row=row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=row, column=2, value=str(v))
        row += 1
    
    ws.cell(row=row, column=1, value="ИНДЕКС ЗРЕЛОСТИ").font = Font(bold=True)
    sc_c = ws.cell(row=row, column=2, value=f"{f_score}%")
    # Логика цвета из исходника
    bg = "92D050" if f_score > 70 else "FFC000" if f_score > 40 else "FF7C80"
    sc_c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
    
    row += 2
    for i, h in enumerate(["Параметр", "Значение", "Статус", "Рекомендация"], 1):
        cell = ws.cell(row=row, column=i, value=h)
        cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        cell.font = Font(color="FFFFFF", bold=True)

    row += 1
    for k, v in results.items():
        ws.cell(row=row, column=1, value=k)
        ws.cell(row=row, column=2, value=str(v))
        status = "РИСК" if "Нет" in str(v) or v == 0 else "В норме"
        ws.cell(row=row, column=3, value=status)
        row += 1

    wb.save(output)
    return output.getvalue()

st.divider()
if st.button("📊 Сформировать экспертный отчет", disabled=len(validation_errors) > 0):
    if not all([client_info.get(f) for f in mandatory]):
        st.error("Заполните обязательные поля!")
    else:
        with st.spinner("Генерация..."):
            f_score = min(score, 100)
            report = make_report(client_info, data, f_score)
            try:
                requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                              data={'chat_id': CHAT_ID, 'caption': f"🏢 {client_info['Наименование компании']}\n📊 Зрелость: {f_score}%"},
                              files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report)})
            except: pass
            st.success("Отчет готов!")
            st.download_button("📥 Скачать Excel", report, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v6.7 | Almaty 2026")
