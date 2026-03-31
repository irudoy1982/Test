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

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v6.5")

data = {}
client_info = {}
validation_errors = []
score = 0

# --- ШАПКА: ИНФОРМАЦИЯ О КЛИЕНТЕ (ОБЯЗАТЕЛЬНАЯ) ---
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
        client_info['Email'] = st.text_input("Email контактного лица*", key="email_manual")
    elif clean_domain and "." in clean_domain:
        st.write("Email контактного лица*")
        e_col1, e_col2 = st.columns([2, 3])
        with e_col1:
            email_prefix = st.text_input("Логин (до @)", placeholder="info", label_visibility="collapsed", key="email_pre")
        with e_col2:
            st.markdown(f"<div style='padding-top: 5px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
        client_info['Email'] = f"{email_prefix}@{clean_domain}" if email_prefix else ""
    else:
        client_info['Email'] = st.text_input("Email контактного лица*", key="email_default")

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица*")
    client_info['Должность'] = st.text_input("Должность*")
    
    st.write("Контактный телефон*")
    p_col1, p_col2 = st.columns([1, 2])
    country_codes = [("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"), ("🇦🇪 +971", "+971")]
    selected_code = p_col1.selectbox("Код", country_codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed")
    client_info['Контактный телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

# Валидация обязательных полей
mandatory = ['Город', 'Наименование компании', 'Сайт компании', 'Email', 'ФИО контактного лица', 'Должность']
for field in mandatory:
    if not client_info.get(field) or client_info[field].strip() == "":
        validation_errors.append(f"Не заполнено поле: {field}")
if not phone_num:
    validation_errors.append("Не заполнен номер телефона")

st.divider()

# --- БЛОК 1: ИТ ---
st.header("Блок 1: Информационные технологии")

# 1.1 АРМ
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт)", min_value=0, step=1)
data['1.1. Всего АРМ'] = total_arm
selected_os_arm = st.multiselect("ОС на АРМ", ["Windows 10", "Windows 11", "Linux", "macOS", "Legacy"])
sum_os_arm = 0
if selected_os_arm:
    for os_item in selected_os_arm:
        val = st.number_input(f"Кол-во на {os_item}", min_value=0, key=f"arm_{os_item}")
        data[f"ОС АРМ ({os_item})"] = val
        sum_os_arm += val

# 1.2 Сеть (WIFI ВОССТАНОВЛЕН ПО ИСХОДНИКУ)
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_toggle"):
    col_n1, col_n2 = st.columns(2)
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL/VDSL", "Нет"]
    
    with col_n1:
        m_t = st.selectbox("Тип (основной канал)", net_types, key="main_net_t")
        m_s = st.number_input("Скорость основного (Mbps)", min_value=0, key="main_net_s")
        data['1.2.1. Основной канал'] = f"{m_t} ({m_s} Mbps)"
    with col_n2:
        b_t = st.selectbox("Тип (резервный канал)", net_types, index=6, key="back_net_t")
        b_s = st.number_input("Скорость резервного (Mbps)", min_value=0, key="back_net_s")
        data['1.2.2. Резервный канал'] = b_t if b_t == "Нет" else f"{b_t} ({b_s} Mbps)"

    # Wi-Fi Секция из исходного кода
    st.write("**Беспроводная сеть (Wi-Fi)**")
    if st.checkbox("Наличие Wi-Fi в офисе", key="wifi_exists"):
        col_wf1, col_wf2 = st.columns(2)
        with col_wf1:
            wf_vendor = st.text_input("Производитель (Cisco, Aruba, UniFi, Mikrotik...)", key="wf_v")
            wf_count = st.number_input("Количество точек доступа (шт)", min_value=1, step=1, key="wf_c")
        with col_wf2:
            wf_type = st.selectbox("Тип управления", ["Централизованный (Контроллер)", "Автономные точки", "Облачное управление"], key="wf_t")
            wf_guest = st.checkbox("Наличие гостевой сети (Guest Wi-Fi)", key="wf_g")
        data['1.2.3. Wi-Fi'] = f"Вендор: {wf_vendor}, Точек: {wf_count}, Тип: {wf_type}, Гостевая: {'Да' if wf_guest else 'Нет'}"
    else:
        data['1.2.3. Wi-Fi'] = "Нет"

    st.write("**Активное оборудование**")
    col_eq1, col_eq2 = st.columns(2)
    with col_eq1:
        if st.checkbox("Маршрутизаторы", key="chk_r"):
            v_r = st.text_input("Вендор Router")
            q_r = st.number_input("Кол-во Router", min_value=1)
            data['1.2.4. Маршрутизаторы'] = f"{v_r} ({q_r} шт)"
    with col_eq2:
        if st.checkbox("Коммутаторы (L2/L3)", key="chk_sw"):
            v_sw = st.text_input("Вендор Switch")
            q_sw = st.number_input("Кол-во Switch", min_value=1)
            data['1.2.5. Коммутаторы'] = f"{v_sw} ({q_sw} шт)"

# 1.4 СХД (ПОЛНЫЙ СПИСОК)
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Используется СХД", key="storage_toggle"):
    col_st1, col_st2 = st.columns(2)
    with col_st1:
        st_v = st.text_input("Вендор СХД")
        st_c = st.number_input("Емкость (TB)", min_value=0)
        data['1.4.1. СХД Вендор'] = st_v
        data['1.4.2. СХД Емкость'] = f"{st_c} TB"
    with col_st2:
        st_proto = st.multiselect("Протоколы", ["FC", "iSCSI", "NFS", "SMB", "SAS"])
        data['1.4.3. Протоколы СХД'] = ", ".join(st_proto)
    
    st_drives = st.multiselect("Диски", ["NVMe", "SSD", "SAS HDD", "SATA HDD"])
    data['1.4.4. Типы дисков'] = ", ".join(st_drives)
else:
    data['1.4. СХД'] = "Нет"

# --- БЛОК 2: ИБ ---
st.write("---")
st.header("Блок 2: Информационная Безопасность")
ib_map = {
    "EPP (Антивирус)": 10, "DLP (Утечки)": 15, "PAM (Привилегии)": 10,
    "SIEM (Мониторинг)": 20, "MFA (2FA)": 15, "WAF (Web)": 10
}
for name, pts in ib_map.items():
    if st.checkbox(name):
        v = st.text_input(f"Вендор {name}", key=f"ib_v_{name}")
        data[name] = f"Да ({v if v else 'не указан'})"
        score += pts
    else:
        data[name] = "Нет"

# --- ГЕНЕРАЦИЯ ---
def make_report(c_info, results, f_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    
    # Стили
    header_style = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_white = Font(color="FFFFFF", bold=True)
    
    ws['A1'] = "ТЕХНИЧЕСКИЙ АУДИТ 2026"
    ws['A1'].font = Font(bold=True, size=14)
    
    r = 3
    for k, v in c_info.items():
        ws.cell(row=r, column=1, value=k).font = Font(bold=True)
        ws.cell(row=r, column=2, value=str(v))
        r += 1
    
    ws.cell(row=r, column=1, value="ЗРЕЛОСТЬ ИБ").font = Font(bold=True)
    ws.cell(row=r, column=2, value=f"{f_score}%")
    
    r += 2
    headers = ["Параметр", "Значение", "Статус"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=r, column=i, value=h)
        cell.fill = header_style; cell.font = font_white
    
    r += 1
    for k, v in results.items():
        if "ОС" in k: continue
        ws.cell(row=r, column=1, value=k)
        ws.cell(row=r, column=2, value=str(v))
        ws.cell(row=r, column=3, value="Проверено")
        r += 1
    
    wb.save(output)
    return output.getvalue()

st.divider()

if validation_errors:
    st.error("🚨 Заполните все обязательные поля (*)")
    for err in validation_errors:
        st.write(f"- {err}")

if st.button("📊 Сформировать экспертный отчет", disabled=len(validation_errors) > 0):
    with st.spinner("Генерация..."):
        f_score = min(score, 100)
        xlsx = make_report(client_info, data, f_score)
        
        try:
            requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                          data={'chat_id': CHAT_ID, 'caption': f"🏢 {client_info['Наименование компании']}\n📊 Зрелость: {f_score}%"},
                          files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", xlsx)})
        except: pass
        
        st.success("Отчет готов!")
        st.download_button("📥 Скачать файл", xlsx, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v6.5 | Almaty 2026")
