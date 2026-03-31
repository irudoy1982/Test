import streamlit as st
import pandas as pd
import os
import requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- 1. CONFIG ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026 Gold", layout="wide", page_icon="🛡️")
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. HEADER ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### 🛡️ Технический аудит: Золотой образ v7.2")
st.divider()

data = {}
client_info = {}
score = 0

# --- БЛОК 0: ОБЩАЯ ИНФОРМАЦИЯ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город:*")
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    site_input = st.text_input("Сайт компании:*", placeholder="example.kz")
    client_info['Сайт компании'] = site_input
    
    # Email Logic
    custom_email = st.checkbox("Email отличается от домена сайта")
    if custom_email:
        client_info['Email'] = st.text_input("Email контактного лица:*")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица:*")
            e_c1, e_c2 = st.columns([1, 2])
            with e_c1:
                prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="epref")
            with e_c2:
                st.markdown(f"<div style='padding-top: 5px; font-weight: bold;'>@{clean_domain}</div>", unsafe_allow_True=True)
            client_info['Email'] = f"{prefix}@{clean_domain}" if prefix else ""
        else: client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    st.write("Контактный телефон:*")
    p_c1, p_c2 = st.columns([1, 2])
    codes = [("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"), ("🇰🇬 +996", "+996"), ("🇹🇯 +992", "+992"), ("🇹🇲 +993", "+993")]
    selected_code = p_c1.selectbox("Код", codes, format_func=lambda x: x[0], label_visibility="collapsed")
    ph_val = p_c2.text_input("Номер", placeholder="707 000 00 00", label_visibility="collapsed")
    client_info['Телефон'] = f"{selected_code[1]} {ph_val}" if ph_val else ""

st.divider()

# --- БЛОК 1: ИТ ИНФРАСТРУКТУРА ---
st.header("Блок 1: Информационные технологии")

# 1.1 АРМ
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)
data['1.1. Всего АРМ'] = total_arm

arm_cols = st.columns(3)
arm_counts = {}
with arm_cols[0]:
    arm_counts['Windows 10'] = st.number_input("Windows 10 (шт):", min_value=0)
    arm_counts['Windows 11'] = st.number_input("Windows 11 (шт):", min_value=0)
with arm_cols[1]:
    arm_counts['Linux'] = st.number_input("Linux Desktop (шт):", min_value=0)
    arm_counts['macOS'] = st.number_input("macOS (шт):", min_value=0)
with arm_cols[2]:
    arm_counts['Legacy (7/8/XP)'] = st.number_input("Legacy (7/8/XP) (шт):", min_value=0)

sum_arm = sum(arm_counts.values())
if sum_arm != total_arm and total_arm > 0:
    st.warning(f"⚠️ Несоответствие: Сумма по ОС ({sum_arm}) не равна общему количеству ({total_arm})")
for k, v in arm_counts.items(): data[f"АРМ: {k}"] = v

# 1.2 Сети
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
col_n1, col_n2 = st.columns(2)
n_list = ["Нет", "Оптика", "Радиорелейная", "Спутник", "4G/5G", "ADSL"]
with col_n1:
    data['1.2.1. Основной канал'] = st.selectbox("Основной канал:", n_list[1:])
    data['Скорость осн.'] = st.number_input("Скорость (Mbit/s):", min_value=0, key="s1")
with col_n2:
    data['1.2.2. Резервный канал'] = st.selectbox("Резервный канал:", n_list, index=0)
    data['Скорость рез.'] = st.number_input("Скорость (Mbit/s):", min_value=0, key="s2")

st.write("**Оборудование:**")
c_net1, c_net2, c_net3 = st.columns(3)
with c_net1:
    if st.checkbox("Ядро сети"): data['Core'] = st.text_input("Вендор Core:")
    if st.checkbox("Маршрутизаторы"): data['Routers'] = st.number_input("Кол-во Routers:", min_value=0)
with c_net2:
    if st.checkbox("Коммутаторы L3"): data['L3 Sw'] = st.number_input("Кол-во L3:", min_value=0)
    if st.checkbox("Коммутаторы L2"): data['L2 Sw'] = st.number_input("Кол-во L2:", min_value=0)
with c_net3:
    if st.checkbox("Wi-Fi Инфраструктура"): data['Wi-Fi'] = st.text_input("Вендор Wi-Fi:")
    if st.checkbox("QoS / Bandwidth"): data['QoS'] = "Да"

if st.checkbox("Межсетевой экран (NGFW)"):
    v_ng = st.text_input("Вендор NGFW:")
    data['1.2.7. NGFW'] = f"Да ({v_ng})"
    score += 20

# 1.3 Серверы (ДЕТАЛИЗИРОВАННЫЕ WINDOWS)
st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
data['1.3.1. Физические серверы'] = st.number_input("Физические серверы (шт):", min_value=0)
data['1.3.2. Виртуальные серверы'] = st.number_input("Виртуальные серверы (шт):", min_value=0)
data['Виртуализация'] = st.selectbox("Платформа:", ["VMware", "Hyper-V", "Proxmox", "KVM/Astra", "Нет"])

st.write("**Детализация ОС Серверов:**")
srv_os1, srv_os2 = st.columns(2)
with srv_os1:
    data['Win Srv 2012/R2'] = st.number_input("Windows Server 2012/R2 (шт):", min_value=0)
    data['Win Srv 2016'] = st.number_input("Windows Server 2016 (шт):", min_value=0)
    data['Win Srv 2019'] = st.number_input("Windows Server 2019 (шт):", min_value=0)
    data['Win Srv 2022'] = st.number_input("Windows Server 2022 (шт):", min_value=0)
with srv_os2:
    data['Linux (RHEL/CentOS)'] = st.number_input("Linux (RHEL/CentOS) (шт):", min_value=0)
    data['Linux (Debian/Ubuntu)'] = st.number_input("Linux (Debian/Ubuntu) (шт):", min_value=0)
    data['Astra Linux'] = st.number_input("Astra Linux (шт):", min_value=0)

# 1.4 СХД
st.write("---")
st.subheader("1.4. СХД")
if st.toggle("Наличие СХД", key="st_t"):
    sc1, sc2 = st.columns(2)
    with sc1:
        st_arch = st.selectbox("Тип:", ["All-Flash (NVMe/SSD)", "Hybrid", "HDD Only"])
        st_v = st.text_input("Вендор СХД:")
        if st.checkbox("HA (2+ контроллера)"): data['СХД HA'] = "Да"
    with sc2:
        st_cap = st.number_input("Емкость (ТБ):", min_value=0)
        data['1.4. СХД'] = f"{st_v} | {st_arch} ({st_cap} TB)"
else: data['1.4. СХД'] = "Нет"

if st.checkbox("Резервное копирование"):
    v_b = st.text_input("Вендор Backup:")
    data["Резервное копирование"] = f"Да ({v_b})"
    score += 20

# 1.5 ИС и WEB (ВОССТАНОВЛЕНО)
st.write("---")
st.subheader("1.5. Информационные системы и Web")
is_col1, is_col2 = st.columns(2)
with is_col1:
    data['ERP/CRM'] = st.text_input("ERP / CRM Системы (1С, SAP, Oracle...):")
    data['Billing'] = st.text_input("Биллинг / Фин. системы:")
with is_col2:
    data['Web External'] = st.text_input("Внешние сайты/порталы:")
    data['Web Internal'] = st.text_input("Внутренние ресурсы (Intranet):")

st.divider()

# --- БЛОК 2: ИБ (ВОССТАНОВЛЕНО) ---
st.header("Блок 2: Информационная Безопасность")
ib_tools = {
    "EPP (Antivirus)": 10, "DLP (Data Loss)": 15, "PAM (Privileged Access)": 10,
    "SIEM (Events)": 20, "VM (Vulnerability)": 10, "EDR/XDR": 15, "WAF (Web App Fire)": 10, "MFA (2FA)": 15
}
ib1, ib2 = st.columns(2)
for i, (name, pts) in enumerate(ib_tools.items()):
    with (ib1 if i < 4 else ib2):
        if st.checkbox(name):
            v_i = st.text_input(f"Производитель {name}:", key=f"ib_{name}")
            data[name] = f"Да ({v_i})"
            score += pts
        else: data[name] = "Нет"

st.divider()

# --- БЛОК 4: РАЗРАБОТКА ---
st.header("Блок 4: Разработка и DevOps")
if st.toggle("Разработка", key="dev_t"):
    d1, d2 = st.columns(2)
    with d1:
        data['4.1. Разработчики'] = st.number_input("Кол-во (чел):", min_value=0)
        data['4.2. Стек'] = st.text_input("Стек (Python/Java...):")
        data['4.3. Git'] = st.selectbox("Git:", ["GitLab", "GitHub", "Bitbucket", "Нет"])
    with d2:
        data['4.4. CI/CD'] = st.selectbox("CI/CD:", ["Jenkins", "GitLab CI", "GitHub Actions", "Нет"])
        data['4.5. Контейнеры'] = st.multiselect("Среды:", ["Docker", "Kubernetes"])
        data['4.6. Среды'] = st.multiselect("Окружения:", ["Dev", "Stage", "Prod"])
        if st.checkbox("SAST/DAST"): data['SAST/DAST'] = "Да"
else: data['4.1. Разработка'] = "Нет"

# --- EXCEL REPORT ---
def generate_report(c_info, results, total_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit_Report"
    
    h_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    w_font = Font(color="FFFFFF", bold=True)
    brd = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    ws.merge_cells('A1:D2'); ws['A1'] = "ТЕХНИЧЕСКИЙ АУДИТ ИНФРАСТРУКТУРЫ 2026"; ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=14)

    row = 4
    for k, v in c_info.items():
        ws.cell(row=row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=row, column=2, value=str(v)); row += 1
    
    ws.cell(row=row, column=1, value="ИНДЕКС ЗРЕЛОСТИ:").font = Font(bold=True)
    ws.cell(row=row, column=2, value=f"{min(total_score, 100)}%"); row += 3

    for i, h in enumerate(["Параметр", "Значение", "Статус", "Рекомендация"], 1):
        c = ws.cell(row=row, column=i, value=h); c.fill = h_fill; c.font = w_font; c.border = brd
    
    row += 1
    n_arm = results.get('1.1. Всего АРМ', 0)
    n_srv = results.get('1.3.1. Физические серверы', 0) + results.get('1.3.2. Виртуальные серверы', 0)
    has_edr = results.get('EDR/XDR') != "Нет"

    for k, v in results.items():
        ws.cell(row=row, column=1, value=k).border = brd
        ws.cell(row=row, column=2, value=str(v)).border = brd
        
        status, rec, color = "В норме", "Ок.", "000000"
        if k == '1.2.2. Резервный канал' and v == "Нет":
            status, rec, color = "КРИТИЧНО", "Единая точка отказа.", "FF0000"
        elif v == "Нет":
            if k == "SIEM (Events)":
                if n_arm < 100 and n_srv < 20 and not has_edr:
                    status, rec, color = "ВНИМАНИЕ", "SIEM не критичен для малого бизнеса.", "FFC000"
                else: status, rec, color = "КРИТИЧНО", "Необходим мониторинг при текущем масштабе.", "FF0000"
            else: status, rec, color = "РИСК", "Рекомендуется внедрение.", "FF0000"

        st_c = ws.cell(row=row, column=3, value=status); st_c.font = Font(color=color, bold=True); st_c.border = brd
        ws.cell(row=row, column=4, value=rec).border = brd
        row += 1

    for c, w in {'A': 35, 'B': 30, 'C': 15, 'D': 55}.items(): ws.column_dimensions[c].width = w
    wb.save(output); return output.getvalue()

# --- FINAL BUTTON ---
st.divider()
is_valid = all([client_info.get('Наименование компании'), client_info.get('Email'), client_info.get('Город'), client_info.get('Телефон')])

if st.button("📊 Сформировать экспертный отчет", disabled=not is_valid):
    report = generate_report(client_info, data, score)
    try:
        requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                      data={"chat_id": CHAT_ID, "caption": f"Аудит: {client_info['Наименование компании']}"}, 
                      files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report)})
    except: pass
    st.success("Отчет сформирован!")
    st.download_button("📥 Скачать экспертный Excel", report, f"Audit_{client_info['Наименование компании']}.xlsx")
elif not is_ready:
    st.warning("⚠️ Заполните все обязательные поля (Компания, Город, Email, Телефон).")
