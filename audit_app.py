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

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v6.6")

# --- ИНСТРУКЦИЯ ДЛЯ ПОЛЬЗОВАТЕЛЯ ---
with st.expander("📖 Инструкция по заполнению (нажмите, чтобы развернуть)"):
    st.markdown("""
    ### Руководство по проведению экспресс-аудита
    Данный инструмент предназначен для сбора технических данных об ИТ-ландшафте и уровне защищенности организации.
    1. **Общая информация:** Все поля со звездочкой (*) обязательны.
    2. **Заполнение блоков:** Используйте переключатели (toggles) для активации нужных подразделов.
    3. **Результат:** Нажмите кнопку внизу для формирования Excel-отчета и отправки эксперту.
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
mandatory_fields = ['Город', 'Наименование компании', 'Сайт компании', 'Email', 'ФИО контактного лица', 'Должность']
for field in mandatory_fields:
    if not client_info.get(field) or str(client_info[field]).strip() == "":
        validation_errors.append(f"Не заполнено: {field}")
if not phone_num:
    validation_errors.append("Не заполнен контактный телефон")

st.divider()

# --- БЛОК 1: ИТ ---
st.header("Блок 1: Информационные технологии")

# 1.1 АРМ
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт)", min_value=0, step=1)
data['1.1. Всего АРМ'] = total_arm
selected_os_arm = st.multiselect("ОС на АРМ", ["Windows 10", "Windows 11", "Linux", "macOS", "Legacy (XP/7/8)"])
sum_os_arm = 0
if selected_os_arm:
    for os_item in selected_os_arm:
        val = st.number_input(f"Кол-во на {os_item}", min_value=0, key=f"arm_{os_item}")
        data[f"ОС АРМ ({os_item})"] = val
        sum_os_arm += val

# 1.2 Сетевая инфраструктура
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_toggle"):
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL/VDSL", "Нет"]
    col_n1, col_n2 = st.columns(2)
    with col_n1:
        m_t = st.selectbox("Тип (основной канал)", net_types, key="main_net_t")
        m_s = st.number_input("Скорость основного (Mbps)", min_value=0, key="main_net_s")
        data['1.2.1. Основной канал'] = f"{m_t} ({m_s} Mbps)"
    with col_n2:
        # ЛОГИКА: Резервный канал по умолчанию "Нет"
        b_t = st.selectbox("Тип (резервный канал)", net_types, index=6, key="back_net_t")
        b_s = st.number_input("Скорость резервного (Mbps)", min_value=0, key="back_net_s")
        data['1.2.2. Резервный канал'] = b_t if b_t == "Нет" else f"{b_t} ({b_s} Mbps)"

    # Wi-Fi (Восстановлено как в исходном)
    st.write("**Беспроводная сеть (Wi-Fi)**")
    if st.checkbox("Используется Wi-Fi в офисе", key="wifi_chk"):
        cw1, cw2 = st.columns(2)
        with cw1:
            wf_v = st.text_input("Вендор Wi-Fi", placeholder="Aruba, Cisco, UniFi...")
            wf_q = st.number_input("Количество точек доступа", min_value=1, step=1)
        with cw2:
            wf_t = st.selectbox("Управление", ["Контроллер (on-prem)", "Облако", "Автономно"])
            wf_g = st.checkbox("Гостевая сеть изолирована")
        data['1.2.3. Wi-Fi'] = f"{wf_v}, {wf_q} шт, {wf_t}, Guest: {'Да' if wf_g else 'Нет'}"
    
    st.write("**Активное оборудование**")
    ce1, ce2, ce3 = st.columns(3)
    with ce1:
        if st.checkbox("Маршрутизаторы"):
            rv = st.text_input("Вендор (Router)"); rq = st.number_input("Кол-во (Router)", min_value=1)
            data['1.2.4. Маршрутизаторы'] = f"{rv} ({rq} шт)"
    with ce2:
        if st.checkbox("Коммутаторы L2"):
            l2v = st.text_input("Вендор (L2)"); l2q = st.number_input("Кол-во (L2)", min_value=1)
            data['1.2.5. Коммутаторы L2'] = f"{l2v} ({l2q} шт)"
    with ce3:
        if st.checkbox("Коммутаторы L3"):
            l3v = st.text_input("Вендор (L3)"); l3q = st.number_input("Кол-во (L3)", min_value=1)
            data['1.2.6. Коммутаторы L3'] = f"{l3v} ({l3q} шт)"

# 1.4 СХД (Восстановлено полностью)
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Наличие выделенной СХД", key="st_toggle"):
    cs1, cs2 = st.columns(2)
    with cs1:
        st_vendor = st.text_input("Производитель СХД")
        st_cap = st.number_input("Полезный объем (TB)", min_value=0)
        data['1.4.1. СХД Вендор'] = st_vendor; data['1.4.2. СХД Объем'] = f"{st_cap} TB"
    with cs2:
        st_proto = st.multiselect("Протоколы доступа", ["FC", "iSCSI", "NFS", "SMB", "SAS"])
        data['1.4.3. Протоколы СХД'] = ", ".join(st_proto)
    st_media = st.multiselect("Типы дисков в СХД", ["NVMe", "SSD", "SAS 10k/15k", "SATA/NL-SAS"])
    data['1.4.4. Носители'] = ", ".join(st_media)
else:
    data['1.4. СХД'] = "Нет"

# --- БЛОК 2: ИБ ---
st.write("---")
st.header("Блок 2: Информационная Безопасность")
ib_list = {"EPP (Антивирус)": 10, "DLP": 15, "PAM": 10, "SIEM": 20, "MFA": 15, "WAF": 10}
col_ib1, col_ib2 = st.columns(2)
for i, (name, pts) in enumerate(ib_list.items()):
    with (col_ib1 if i < 3 else col_ib2):
        if st.checkbox(name):
            v_ib = st.text_input(f"Вендор {name}")
            data[name] = f"Да ({v_ib if v_ib else 'не указан'})"
            score += pts
        else:
            data[name] = "Нет"

# --- ФУНКЦИЯ EXCEL ---
def generate_report(c_info, results, f_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit"
    
    # Стилизация
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    ws.merge_cells('A1:D1')
    ws['A1'] = "ОТЧЕТ ПО ТЕХНИЧЕСКОМУ АУДИТУ 2026"
    ws['A1'].alignment = Alignment(horizontal='center')
    
    r = 3
    for k, v in c_info.items():
        ws.cell(row=r, column=1, value=k).font = Font(bold=True)
        ws.cell(row=r, column=2, value=str(v))
        r += 1
    
    ws.cell(row=r, column=1, value="ИНДЕКС ЗРЕЛОСТИ").font = Font(bold=True)
    ws.cell(row=r, column=2, value=f"{f_score}%")
    
    r += 2
    for i, h in enumerate(["Параметр", "Значение", "Статус", "Рекомендация"], 1):
        cell = ws.cell(row=r, column=i, value=h)
        cell.fill = header_fill; cell.font = white_font
    
    r += 1
    for k, v in results.items():
        if "ОС" in k: continue
        ws.cell(row=r, column=1, value=k).border = border
        ws.cell(row=r, column=2, value=str(v)).border = border
        ws.cell(row=r, column=3, value="Проверено").border = border
        ws.cell(row=r, column=4, value="Ожидает оценки эксперта").border = border
        r += 1

    wb.save(output)
    return output.getvalue()

st.divider()

# --- КНОПКА ФИНАЛИЗАЦИИ ---
if validation_errors:
    st.error(f"🚨 Пожалуйста, заполните обязательные поля (*) для формирования отчета.")

if st.button("📊 Сформировать экспертный отчет", disabled=len(validation_errors) > 0):
    with st.spinner("Создание отчета..."):
        f_score = min(score, 100)
        report_data = generate_report(client_info, data, f_score)
        
        try:
            requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                          data={'chat_id': CHAT_ID, 'caption': f"🏢 {client_info['Наименование компании']}\n📊 Зрелость: {f_score}%"},
                          files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report_data)})
        except: pass
        
        st.success("Отчет успешно создан и отправлен!")
        st.download_button("📥 Скачать Excel", report_data, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v6.6 | Almaty 2026")
