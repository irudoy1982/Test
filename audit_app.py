import streamlit as st
import pandas as pd
import os
import requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- 1. КОНФИГУРАЦИЯ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ПРИВЕТСТВИЕ ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### Полный технический аудит инфраструктуры (v7.1)")
st.divider()

data = {}
client_info = {}
score = 0

# --- БЛОК: ОБЩАЯ ИНФОРМАЦИЯ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город:*")
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    site_input = st.text_input("Сайт компании:*", placeholder="example.kz")
    client_info['Сайт компании'] = site_input
    
    # СТРОГАЯ ВАЛИДАЦИЯ EMAIL
    custom_email = st.checkbox("Email отличается от домена сайта")
    if custom_email:
        client_info['Email'] = st.text_input("Email контактного лица:*", key="manual_email")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица:*")
            e_c1, e_c2 = st.columns([1, 2])
            with e_c1:
                prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="email_prefix")
            with e_c2:
                st.markdown(f"<div style='padding-top: 5px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
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

# 1.1 АРМ (ДЕТАЛЬНО)
st.subheader("1.1. Конечные точки (АРМ)")
c_arm1, c_arm2 = st.columns(2)
with c_arm1:
    total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)
    data['1.1. Всего АРМ'] = total_arm
with c_arm2:
    selected_os_arm = st.multiselect("Используемые ОС на АРМ:", ["Windows 10", "Windows 11", "Linux", "macOS", "Legacy (XP/7/8)"])
    for os_item in selected_os_arm:
        data[f"ОС АРМ ({os_item})"] = st.number_input(f"Кол-во на {os_item}:", min_value=0, step=1, key=f"arm_{os_item}")

# 1.2 СЕТИ (ДЕТАЛЬНО)
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Собственная сеть", key="net_toggle", value=True):
    net_types = ["Нет", "Оптика", "Радиорелейная", "Спутник", "4G/5G/Starlink", "ADSL"]
    col_n1, col_n2 = st.columns(2)
    with col_n1:
        data['1.2.1. Основной канал'] = st.selectbox("Тип осн. канала:", net_types[1:])
        data['main_speed'] = st.number_input("Скорость осн. канала (Mbit/s):", min_value=0)
    with col_n2:
        data['1.2.2. Резервный канал'] = st.selectbox("Тип рез. канала:", net_types, index=0)
        data['back_speed'] = st.number_input("Скорость рез. канала (Mbit/s):", min_value=0)

    st.write("**Активное оборудование:**")
    nc1, nc2, nc3 = st.columns(3)
    with nc1:
        if st.checkbox("Ядро (Core Switch)"): data['Core'] = st.text_input("Вендор Core:")
        if st.checkbox("Маршрутизаторы"): data['Routers'] = st.number_input("Кол-во Routers:", min_value=0)
    with nc2:
        if st.checkbox("Коммутаторы L3"): data['L3 Switches'] = st.number_input("Кол-во L3:", min_value=0)
        if st.checkbox("Коммутаторы L2"): data['L2 Switches'] = st.number_input("Кол-во L2:", min_value=0)
    with nc3:
        if st.checkbox("Wi-Fi (Контроллер + Точки)"): data['Wi-Fi'] = st.text_input("Вендор Wi-Fi:")
        if st.checkbox("QoS / Bandwidth Manager"): data['QoS'] = "Да"

    if st.checkbox("Межсетевой экран (NGFW)"):
        v_ng = st.text_input("Вендор NGFW (Fortigate/CheckPoint/PaloAlto):")
        data['1.2.7. NGFW'] = f"Да ({v_ng})"
        score += 20

# 1.3 СЕРВЕРЫ (ДЕТАЛЬНО)
st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
sc1, sc2 = st.columns(2)
with sc1:
    data['1.3.1. Физические серверы'] = st.number_input("Кол-во физ. серверов (шт):", min_value=0)
    data['Виртуализация'] = st.selectbox("Платформа:", ["VMware", "Hyper-V", "Proxmox", "KVM/Astra/R-Virtualization", "Нет"])
with sc2:
    data['1.3.2. Виртуальные серверы'] = st.number_input("Кол-во вирт. серверов (шт):", min_value=0)
    data['ОС Серверов'] = st.multiselect("Операционные системы:", 
                                         ["Win Server 2012/R2", "Win Server 2016/19/22", "Linux (Ubuntu/CentOS/RHEL)", "Astra Linux"])

# 1.4 СХД (ДЕТАЛЬНО)
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Наличие выделенной СХД", key="storage_toggle"):
    stc1, stc2 = st.columns(2)
    with stc1:
        st_arch = st.selectbox("Архитектура массива:", ["All-Flash (NVMe/SSD)", "Hybrid (SSD+HDD)", "HDD Only"])
        st_vendor = st.text_input("Производитель СХД (Dell, Huawei, HPE, NetApp):")
        if st.checkbox("Дублирование контроллеров (HA)"): data['СХД HA'] = "Да"
    with stc2:
        st_cap = st.number_input("Полезная емкость (ТБ):", min_value=0)
        st_prot = st.multiselect("Протоколы доступа:", ["FC (Fibre Channel)", "iSCSI", "NFS", "SMB/S3", "NVMe-oF"])
    data['1.4. СХД'] = f"{st_vendor} | {st_arch} ({st_cap} TB)"
else:
    data['1.4. СХД'] = "Нет"

if st.checkbox("Система резервного копирования (Backup)"):
    v_backup = st.text_input("Производитель ПО (Veeam, Кибербекап, Veritas):")
    data["Резервное копирование"] = f"Да ({v_backup})"
    score += 20

st.divider()

# --- БЛОК 2: ИНФОРМАЦИОННАЯ БЕЗОПАСНОСТЬ (ПОЛНЫЙ СПИСОК) ---
st.header("Блок 2: Информационная Безопасность")
ib_tools = {
    "EPP (Антивирусная защита)": 10, "DLP (Защита от утечек)": 15, 
    "PAM (Управление доступом)": 10, "SIEM (Мониторинг событий)": 20, 
    "VM (Сканер уязвимостей)": 10, "EDR/XDR": 15, 
    "WAF (Защита Web-ресурсов)": 10, "MFA (2FA)": 15
}
col_ib1, col_ib2 = st.columns(2)
for i, (name, pts) in enumerate(ib_tools.items()):
    with (col_ib1 if i < 4 else col_ib2):
        if st.checkbox(name):
            v_ib = st.text_input(f"Производитель {name}:", key=f"v_{name}")
            data[name] = f"Да ({v_ib})"
            score += pts
        else:
            data[name] = "Нет"

st.divider()

# --- БЛОК 4: РАЗРАБОТКА И DEVOPS (ПОЛНЫЙ ЦИКЛ) ---
st.header("Блок 4: Разработка и DevOps")
if st.toggle("Внутренняя разработка", key="dev_toggle"):
    dc1, dc2 = st.columns(2)
    with dc1:
        data['4.1. Штат разработки'] = st.number_input("Кол-во разработчиков (чел):", min_value=0)
        data['4.2. Стек'] = st.text_input("Стек (Python, Java, PHP, JS...):")
        data['4.3. Хранение кода'] = st.selectbox("Репозиторий:", ["GitLab", "GitHub", "Bitbucket", "Локальный Git", "Нет"])
        if st.checkbox("Процесс Code Review"): data['Code Review'] = "Да"
    with dc2:
        data['4.4. CI/CD'] = st.selectbox("Автоматизация:", ["Jenkins", "GitLab CI", "GitHub Actions", "TeamCity", "Нет"])
        data['4.5. Контейнеризация'] = st.multiselect("Инструменты:", ["Docker", "Kubernetes (K8s)", "OpenShift"])
        data['4.6. Среды'] = st.multiselect("Наличие окружений:", ["Development", "Staging/Testing", "Production"])
        if st.checkbox("Анализ безопасности кода (SAST/DAST)"): data['SAST/DAST'] = "Да"
else:
    data['4.1. Разработка'] = "Нет"

# --- ГЕНЕРАЦИЯ ОТЧЕТА ---
def generate_excel(c_info, results, total_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit_Report"
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    ws.merge_cells('A1:D2'); ws['A1'] = "ЭКСПЕРТНЫЙ ТЕХНИЧЕСКИЙ АУДИТ 2026"; ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=14)

    row = 4
    for k, v in c_info.items():
        ws.cell(row=row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=row, column=2, value=str(v)); row += 1
    
    ws.cell(row=row, column=1, value="ИНДЕКС ЗРЕЛОСТИ ИТ/ИБ:").font = Font(bold=True)
    ws.cell(row=row, column=2, value=f"{min(total_score, 100)}%"); row += 3

    headers = ["Параметр", "Значение", "Статус", "Рекомендация / Обоснование"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=i, value=h); cell.fill = header_fill; cell.font = white_font; cell.border = border
    
    row += 1
    n_arm = results.get('1.1. Всего АРМ', 0)
    n_srv = results.get('1.3.1. Физические серверы', 0) + results.get('1.3.2. Виртуальные серверы', 0)
    has_edr = results.get('EDR/XDR') != "Нет"

    for k, v in results.items():
        if any(x in k for x in ["ОС", "speed", "Виртуализация"]): continue
        ws.cell(row=row, column=1, value=k).border = border
        ws.cell(row=row, column=2, value=str(v)).border = border
        
        status, rec, color = "В норме", "Поддерживать текущее состояние.", "000000"

        if k == '1.2.2. Резервный канал' and v == "Нет":
            status, rec, color = "КРИТИЧНО", "Высокий риск остановки бизнеса при аварии на линии.", "FF0000"
        elif v == "Нет":
            if k == "SIEM (Мониторинг событий)":
                if n_arm < 100 and n_srv < 20 and not has_edr:
                    status, rec, color = "ВНИМАНИЕ", "SIEM не критичен для текущего масштаба. Рекомендован при росте.", "FFC000"
                else: status, rec, color = "КРИТИЧНО", "Необходим централизованный сбор событий при текущем масштабе.", "FF0000"
            elif k == "1.4. СХД" and n_srv > 10:
                status, rec, color = "КРИТИЧНО", "Риск простоя из-за отсутствия СХД (High Availability).", "FF0000"
            else:
                status, rec, color = "РИСК", "Рекомендуется внедрение для минимизации угроз.", "FF0000"

        st_c = ws.cell(row=row, column=3, value=status); st_c.font = Font(color=color, bold=True); st_c.border = border
        ws.cell(row=row, column=4, value=rec).border = border
        row += 1

    for c, w in {'A': 35, 'B': 30, 'C': 15, 'D': 55}.items(): ws.column_dimensions[c].width = w
    wb.save(output); return output.getvalue()

# --- ФИНАЛ ---
st.divider()
is_ready = all([client_info.get('Наименование компании'), client_info.get('Email'), client_info.get('Город'), client_info.get('Телефон')])

if st.button("📊 Сформировать экспертный отчет", disabled=not is_ready):
    report = generate_excel(client_info, data, score)
    try:
        requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                      data={"chat_id": CHAT_ID, "caption": f"Новый аудит: {client_info['Наименование компании']}"}, 
                      files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report)})
    except: pass
    st.success("Отчет сформирован!")
    st.download_button("📥 Скачать экспертный Excel", report, f"Audit_{client_info['Наименование компании']}.xlsx")
elif not is_ready:
    st.warning("⚠️ Для активации кнопки заполните обязательные поля: Город, Компания, Email и Телефон.")
