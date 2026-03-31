import streamlit as st
import pandas as pd
import os
import requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

# --- НАСТРОЙКИ TELEGRAM (ИЗ SECRETS) ---
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ВИЗУАЛ ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### Экспертная оценка технологического стека")
st.divider()

st.title("📋 Опросник: Технический аудит ИТ и ИБ v7.0")

data = {}
client_info = {}
score = 0

# --- ШАПКА: ОБЩАЯ ИНФОРМАЦИЯ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город:*")
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    site_input = st.text_input("Сайт компании:*", placeholder="example.kz")
    client_info['Сайт компании'] = site_input
    
    custom_email = st.checkbox("Email отличается от домена сайта")
    if custom_email:
        client_info['Email'] = st.text_input("Email контактного лица:*")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица:*")
            e_c1, e_c2 = st.columns([1, 2])
            with e_c1:
                prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="mail_p")
            with e_c2:
                st.markdown(f"<div style='padding-top: 5px; font-weight: bold;'>@{clean_domain}</div>", unsafe_allow_html=True)
            client_info['Email'] = f"{prefix}@{clean_domain}" if prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    
    st.write("Контактный телефон:*")
    p_c1, p_c2 = st.columns([1, 2])
    phone_codes = [
        ("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"), 
        ("🇰🇬 +996", "+996"), ("🇹🇯 +992", "+992"), ("🇹🇲 +993", "+993")
    ]
    selected_code = p_c1.selectbox("Код", phone_codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_num = p_c2.text_input("Номер", placeholder="707 000 00 00", label_visibility="collapsed")
    client_info['Телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

st.divider()

# --- БЛОК 1: ИТ ИНФРАСТРУКТУРА ---
st.header("Блок 1: Информационные технологии")

# 1.1 АРМ
st.subheader("1.1. Конечные точки (АРМ)")
data['1.1. Всего АРМ'] = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)

# 1.2 Сети (ПОЛНОЕ ВОССТАНОВЛЕНИЕ)
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сеть", key="net_sw", value=True):
    n_types = ["Нет", "Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL"]
    col_n1, col_n2 = st.columns(2)
    with col_n1:
        data['1.2.1. Основной канал'] = st.selectbox("Тип осн. канала:", n_types[1:])
        data['1.2.1. Скорость (Mbit)'] = st.number_input("Скорость осн. канала:", min_value=0)
    with col_n2:
        data['1.2.2. Резервный канал'] = st.selectbox("Тип рез. канала:", n_types, index=0)
        data['1.2.2. Скорость (Mbit)'] = st.number_input("Скорость рез. канала:", min_value=0)

    st.write("**Активное оборудование:**")
    nc1, nc2, nc3 = st.columns(3)
    with nc1:
        if st.checkbox("Ядро (Core Switch)"): data['Core'] = st.text_input("Вендор Core:")
        if st.checkbox("Маршрутизаторы"): data['Routers'] = st.number_input("Кол-во Routers:", min_value=0)
    with nc2:
        if st.checkbox("L3 Коммутаторы"): data['L3 Sw'] = st.number_input("Кол-во L3 Sw:", min_value=0)
        if st.checkbox("L2 Коммутаторы"): data['L2 Sw'] = st.number_input("Кол-во L2 Sw:", min_value=0)
    with nc3:
        if st.checkbox("Wi-Fi (AP)"): data['Wi-Fi'] = st.text_input("Вендор Wi-Fi:")
        if st.checkbox("QoS (Управление трафиком)"): data['QoS'] = "Да"

    if st.checkbox("Межсетевой экран (NGFW)"):
        v_ng = st.text_input("Вендор NGFW (Fortigate/CheckPoint/PaloAlto):")
        data['1.2.7. NGFW'] = f"Да ({v_ng})"
        score += 20

# 1.3 Серверы (ПОЛНОЕ ВОССТАНОВЛЕНИЕ)
st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
sc1, sc2 = st.columns(2)
with sc1:
    data['1.3.1. Физические серверы'] = st.number_input("Кол-во физ. серверов:", min_value=0)
    data['Виртуализация'] = st.selectbox("Платформа:", ["VMware", "Hyper-V", "Proxmox", "KVM/Astra", "Нет"])
with sc2:
    data['1.3.2. Виртуальные серверы'] = st.number_input("Кол-во вирт. серверов:", min_value=0)
    data['ОС Серверов'] = st.multiselect("Операционные системы:", 
                                         ["Windows Server 2012/R2", "Windows Server 2016/19/22", "Linux (RHEL/Ubuntu/CentOS)", "Astra Linux"])

# 1.4 СХД (ПОЛНОЕ ВОССТАНОВЛЕНИЕ)
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Наличие СХД", key="st_sw"):
    stc1, stc2 = st.columns(2)
    with stc1:
        st_a = st.selectbox("Тип массива:", ["All-Flash (NVMe/SSD)", "Hybrid (Flash+HDD)", "HDD Only"])
        st_v = st.text_input("Вендор СХД (Dell/Huawei/HPE/NetApp):")
        if st.checkbox("HA (Отказоустойчивые контроллеры)"): data['СХД HA'] = "Да"
    with stc2:
        st_p = st.multiselect("Протоколы доступа:", ["FC", "iSCSI", "NFS", "SMB/S3", "NVMe-oF"])
        st_cap = st.number_input("Полезная емкость (ТБ):", min_value=0)
    data['1.4. СХД'] = f"{st_v} | {st_a} ({st_cap} TB)"
else:
    data['1.4. СХД'] = "Нет"

if st.checkbox("Резервное копирование (Backup)"):
    v_b = st.text_input("Вендор ПО Backup (Veeam/Кибербекап/Veritas):")
    data["Резервное копирование"] = f"Да ({v_b})"
    score += 20

st.divider()

# --- БЛОК 2: ИБ ---
st.header("Блок 2: Информационная Безопасность")
ib_map = {
    "EPP (Антивирус)": 10, "DLP (Защита от утечек)": 15, "PAM (Управление привилегиями)": 10,
    "SIEM (Мониторинг событий)": 20, "VM (Сканер уязвимостей)": 10, "EDR/XDR": 15, "MFA (2FA)": 15
}
ib_col1, ib_col2 = st.columns(2)
for i, (name, pts) in enumerate(ib_map.items()):
    with (ib_col1 if i < 4 else ib_col2):
        if st.checkbox(name):
            v_i = st.text_input(f"Вендор {name}:", key=f"ib_{name}")
            data[name] = f"Да ({v_i})"
            score += pts
        else: data[name] = "Нет"

st.divider()

# --- БЛОК 4: РАЗРАБОТКА (ПОЛНОЕ ВОССТАНОВЛЕНИЕ) ---
st.header("Блок 4: Разработка и DevOps")
if st.toggle("Внутренняя разработка", key="dev_sw"):
    dc1, dc2 = st.columns(2)
    with dc1:
        data['4.1. Разработчики'] = st.number_input("Кол-во в штате:", min_value=0)
        data['4.2. Стек'] = st.text_input("Языки/Фреймворки (Python/Java/PHP):")
        data['4.3. Репозиторий'] = st.selectbox("Хранение кода:", ["GitLab", "GitHub", "Bitbucket", "Нет"])
        if st.checkbox("Процесс Code Review"): data['Code Review'] = "Да"
    with dc2:
        data['4.4. CI/CD'] = st.selectbox("Инструмент CI/CD:", ["Jenkins", "GitLab CI", "GitHub Actions", "TeamCity", "Нет"])
        data['4.5. Контейнеры'] = st.multiselect("Среды исполнения:", ["Docker", "Kubernetes (K8s)", "OpenShift"])
        data['4.6. Окружения'] = st.multiselect("Среды:", ["Development", "Staging", "Production"])
        if st.checkbox("Использование SAST/DAST"): data['SAST/DAST'] = "Да"
else:
    data['4.1. Разработка'] = "Нет"

# --- ГЕНЕРАЦИЯ EXCEL (С ЛОГИКОЙ SIEM И РЕЗЕРВА) ---
def build_report(c_info, results, total_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit_2026"
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    ws.merge_cells('A1:D2'); ws['A1'] = "ЭКСПЕРТНЫЙ ТЕХНИЧЕСКИЙ АУДИТ ИНФРАСТРУКТУРЫ"; ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=14)

    row = 4
    for k, v in c_info.items():
        ws.cell(row=row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=row, column=2, value=str(v)); row += 1
    
    ws.cell(row=row, column=1, value="ИНДЕКС ЗРЕЛОСТИ:").font = Font(bold=True)
    ws.cell(row=row, column=2, value=f"{min(total_score, 100)}%"); row += 3

    for i, h in enumerate(["Параметр", "Значение", "Статус", "Рекомендация"], 1):
        cell = ws.cell(row=row, column=i, value=h); cell.fill = header_fill; cell.font = white_font; cell.border = border
    
    row += 1
    n_arm = results.get('1.1. Всего АРМ', 0)
    n_srv = results.get('1.3.1. Физические серверы', 0) + results.get('1.3.2. Виртуальные серверы', 0)
    has_edr = results.get('EDR/XDR') != "Нет"

    for k, v in results.items():
        if any(x in k for x in ["ОС", "Скорость", "Виртуализация"]): continue
        ws.cell(row=row, column=1, value=k).border = border
        ws.cell(row=row, column=2, value=str(v)).border = border
        
        status, rec, color = "В норме", "Ок.", "000000"
        if k == '1.2.2. Резервный канал' and v == "Нет":
            status, rec, color = "КРИТИЧНО", "Единая точка отказа. Требуется внедрение резервного канала связи.", "FF0000"
        elif v == "Нет":
            if k == "SIEM (Мониторинг событий)" and n_arm < 100 and n_srv < 20 and not has_edr:
                status, rec, color = "ВНИМАНИЕ", "SIEM не критичен при текущем масштабе. Рекомендован при росте.", "FFC000"
            else:
                status, rec, color = "РИСК", "Рекомендуется внедрение для обеспечения безопасности.", "FF0000"

        st_c = ws.cell(row=row, column=3, value=status); st_c.font = Font(color=color, bold=True); st_c.border = border
        ws.cell(row=row, column=4, value=rec).border = border
        row += 1

    for c, w in {'A': 35, 'B': 30, 'C': 15, 'D': 55}.items(): ws.column_dimensions[c].width = w
    wb.save(output); return output.getvalue()

# --- ФИНАЛЬНАЯ КНОПКА ---
st.divider()
is_ready = all([client_info.get('Наименование компании'), client_info.get('Email'), client_info.get('Город'), client_info.get('Телефон')])

if st.button("📊 Сформировать экспертный отчет", disabled=not is_ready):
    report_file = build_report(client_info, data, score)
    try:
        requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                      data={"chat_id": CHAT_ID, "caption": f"Новый аудит: {client_info['Наименование компании']}"}, 
                      files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report_file)})
    except: pass
    st.success("Отчет успешно создан!")
    st.download_button("📥 Скачать экспертный Excel", report_file, f"Audit_{client_info['Наименование компании']}.xlsx")
elif not is_ready:
    st.warning("⚠️ Заполните все обязательные поля (Город, Компания, Email и Телефон) для генерации отчета.")
