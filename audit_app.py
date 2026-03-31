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
    country_codes = [("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"), ("🇦🇪 +971", "+971"), ("🇹🇷 +90", "+90")]
    selected_code = p_col1.selectbox("Код", country_codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed")
    client_info['Контактный телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

# Валидация обязательных полей
mandatory = ['Город', 'Наименование компании', 'Сайт компании', 'Email', 'ФИО контактного лица', 'Должность']
for field in mandatory:
    if not client_info.get(field) or client_info[field].strip() == "":
        validation_errors.append(f"Поле '{field}' не заполнено")
if not phone_num:
    validation_errors.append("Контактный телефон не заполнен")

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
        val = st.number_input(f"Кол-во на {os_item}", min_value=0, step=1, key=f"arm_{os_item}")
        data[f"ОС АРМ ({os_item})"] = val
        sum_os_arm += val
if total_arm > 0 and sum_os_arm != total_arm:
    st.warning(f"⚠️ Ошибка: Сумма ОС ({sum_os_arm}) != Всего АРМ ({total_arm})")
    validation_errors.append("Ошибка баланса АРМ")

# 1.2 Сеть (ПОЛНЫЙ СПИСОК)
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_toggle"):
    col_n1, col_n2 = st.columns(2)
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL/VDSL", "Нет"]
    
    with col_n1:
        m_t = st.selectbox("Тип (основной канал)", net_types, key="main_net_t")
        m_s = st.number_input("Скорость основного (Mbit/s)", min_value=0, key="main_net_s")
        data['1.2.1. Основной канал'] = f"{m_t} ({m_s} Mbps)"
    with col_n2:
        # Резервный канал по умолчанию "Нет" (index=6)
        b_t = st.selectbox("Тип (резервный канал)", net_types, index=6, key="back_net_t")
        b_s = st.number_input("Скорость резервного (Mbit/s)", min_value=0, key="back_net_s")
        data['1.2.2. Резервный канал'] = b_t if b_t == "Нет" else f"{b_t} ({b_s} Mbps)"

    st.write("**Активное сетевое оборудование**")
    c_eq1, c_eq2, c_eq3 = st.columns(3)
    with c_eq1:
        if st.checkbox("Маршрутизаторы"):
            v = st.text_input("Вендор (Router)", key="v_r")
            q = st.number_input("Кол-во (Router)", min_value=1, key="q_r")
            data['1.2.3. Маршрутизаторы'] = f"{v} ({q} шт)"
    with c_eq2:
        if st.checkbox("Коммутаторы L2"):
            v = st.text_input("Вендор (L2)", key="v_l2")
            q = st.number_input("Кол-во (L2)", min_value=1, key="q_l2")
            data['1.2.4. Коммутаторы L2'] = f"{v} ({q} шт)"
    with c_eq3:
        if st.checkbox("Коммутаторы L3"):
            v = st.text_input("Вендор (L3)", key="v_l3")
            q = st.number_input("Кол-во (L3)", min_value=1, key="q_l3")
            data['1.2.5. Коммутаторы L3'] = f"{v} ({q} шт)"

    if st.checkbox("Межсетевой экран (NGFW)", key="chk_ngfw"):
        v_ngfw = st.text_input("Производитель NGFW")
        data['1.2.6. NGFW'] = f"Да ({v_ngfw if v_ngfw else 'не указан'})"
        score += 20
    else:
        data['1.2.6. NGFW'] = "Нет"

# 1.3 Серверы
st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
col_srv1, col_srv2 = st.columns(2)
with col_srv1:
    phys = st.number_input("Физические серверы (шт)", min_value=0)
    data['1.3.1. Физические серверы'] = phys
with col_srv2:
    virt = st.number_input("Виртуальные серверы (шт)", min_value=0)
    data['1.3.2. Виртуальные серверы'] = virt

if st.checkbox("Резервное копирование"):
    v_b = st.text_input("Вендор бэкапа")
    data['1.3.3. Резервное копирование'] = f"Да ({v_b})"
    score += 20
else:
    data['1.3.3. Резервное копирование'] = "Нет"

# 1.4 СХД (ПОЛНЫЙ СПИСОК)
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Есть выделенная СХД", key="storage_toggle"):
    col_st1, col_st2 = st.columns(2)
    with col_st1:
        st_v = st.text_input("Производитель СХД (HP, Dell, Huawei и др.)")
        st_c = st.number_input("Полезная емкость (TB)", min_value=0)
        data['1.4.1. СХД Производитель'] = st_v if st_v else "не указан"
        data['1.4.2. СХД Емкость'] = f"{st_c} TB"
    with col_st2:
        st_conn = st.multiselect("Протоколы доступа", ["FC (Fiber Channel)", "iSCSI", "NFS/SMB", "SAS"])
        data['1.4.3. Протоколы СХД'] = ", ".join(st_conn) if st_conn else "не указаны"
    
    st_media = st.multiselect("Типы дисков", ["All-Flash (SSD/NVMe)", "Hybrid", "HDD Only"])
    data['1.4.4. Типы носителей'] = ", ".join(st_media) if st_media else "не указаны"
else:
    data['1.4. СХД'] = "Нет"

# --- БЛОК 2: ИБ ---
st.write("---")
st.header("Блок 2: Информационная Безопасность")
ib_tools = {
    "EPP (Антивирус)": 10, "DLP (Утечки)": 15, "PAM (Привилегии)": 10,
    "SIEM (Мониторинг)": 20, "VM (Уязвимости)": 10, "EDR/XDR": 15,
    "MFA (Аутентификация)": 15, "WAF (Защита Web)": 10
}
col_ib1, col_ib2 = st.columns(2)
for i, (name, pts) in enumerate(ib_tools.items()):
    with (col_ib1 if i < 4 else col_ib2):
        if st.checkbox(name):
            v_ib = st.text_input(f"Вендор {name}", key=f"v_{name}")
            data[name] = f"Да ({v_ib if v_ib else 'не указан'})"
            score += pts
        else:
            data[name] = "Нет"

# --- ГЕНЕРАЦИЯ ОТЧЕТА ---
def generate_excel(c_info, results, f_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit_Report"
    
    # Стили
    blue_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Заголовок
    ws.merge_cells('A1:D1')
    ws['A1'] = "ТЕХНИЧЕСКИЙ АУДИТ 2026 - ЭКСПЕРТНЫЙ ОТЧЕТ"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center')

    # Данные клиента
    row = 3
    for k, v in c_info.items():
        ws.cell(row=row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=row, column=2, value=str(v))
        row += 1
    
    ws.cell(row=row, column=1, value="ИНДЕКС ЗРЕЛОСТИ").font = Font(bold=True)
    ws.cell(row=row, column=2, value=f"{f_score}%").font = Font(bold=True)

    # Таблица результатов
    row += 2
    headers = ["Параметр", "Значение", "Статус", "Рекомендация"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=i, value=h)
        cell.fill = blue_fill
        cell.font = white_font
    
    row += 1
    for k, v in results.items():
        if "ОС" in k: continue # Скрываем детальную разбивку ОС в таблице
        ws.cell(row=row, column=1, value=k).border = border
        ws.cell(row=row, column=2, value=str(v)).border = border
        
        status = "В норме"
        rec = "Соответствует базовым требованиям"
        if v == "Нет" or "0 TB" in str(v):
            status = "РИСК"
            rec = "Требуется анализ необходимости внедрения"
        
        ws.cell(row=row, column=3, value=status).border = border
        ws.cell(row=row, column=4, value=rec).border = border
        row += 1

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['D'].width = 50
    wb.save(output)
    return output.getvalue()

st.divider()

# ФИНАЛЬНАЯ КНОПКА
if validation_errors:
    st.error("🚨 Пожалуйста, исправьте ошибки для формирования отчета:")
    for err in validation_errors:
        st.write(f"- {err}")

if st.button("📊 Сформировать экспертный отчет", disabled=len(validation_errors) > 0):
    with st.spinner("Создание отчета..."):
        final_score = min(score, 100)
        excel_data = generate_excel(client_info, data, final_score)
        
        # Отправка в TG
        try:
            requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                          data={'chat_id': CHAT_ID, 'caption': f"🏢 {client_info['Наименование компании']}\n📊 Зрелость: {final_score}%"},
                          files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", excel_data)})
        except: pass
        
        st.success("Отчет успешно сформирован!")
        st.download_button("📥 Скачать Excel", excel_data, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v6.3 | Almaty 2026")
