import streamlit as st
import pandas as pd
import os
import requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

# --- НАСТРОЙКИ TELEGRAM ---
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ЛОГОТИП ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### Профессиональный аудит инфраструктуры")
st.divider()

st.title("📋 Опросник: Технический аудит ИТ и ИБ v6.9")

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
    
    # ЛОГИКА EMAIL (ОБЯЗАТЕЛЬНО)
    custom_email = st.checkbox("Email отличается от домена сайта")
    if custom_email:
        client_info['Email'] = st.text_input("Email контактного лица:*")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица:*")
            e_c1, e_c2 = st.columns([1, 2])
            with e_c1:
                prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="email_p")
            with e_c2:
                st.markdown(f"<div style='padding-top: 5px; font-weight: bold;'>@{clean_domain}</div>", unsafe_allow_html=True)
            client_info['Email'] = f"{prefix}@{clean_domain}" if prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    
    # ТЕЛЕФОН С КОДАМИ
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

# --- БЛОК 1: ИТ ---
st.header("Блок 1: Информационные технологии")

# 1.1 АРМ
st.subheader("1.1. Конечные точки (АРМ)")
data['1.1. Всего АРМ'] = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)

# 1.2 Сетевая инфраструктура (ВОССТАНОВЛЕНО ВСЁ)
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сеть", key="net_sw", value=True):
    n_types = ["Нет", "Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL"]
    col_n1, col_n2 = st.columns(2)
    with col_n1:
        data['1.2.1. Основной канал'] = st.selectbox("Основной канал:", n_types[1:])
    with col_n2:
        data['1.2.2. Резервный канал'] = st.selectbox("Резервный канал:", n_types, index=0)

    st.write("**Оборудование и сервисы:**")
    nc1, nc2, nc3 = st.columns(3)
    with nc1:
        if st.checkbox("Ядро (Core)"): data['Core Switch'] = st.text_input("Вендор Core:")
        if st.checkbox("Маршрутизаторы"): data['Routers'] = st.number_input("Кол-во:", min_value=0, key="r_cnt")
    with nc2:
        if st.checkbox("L3 Коммутаторы"): data['L3 Switches'] = st.number_input("Кол-во L3:", min_value=0)
        if st.checkbox("L2 Коммутаторы"): data['L2 Switches'] = st.number_input("Кол-во L2:", min_value=0)
    with nc3:
        if st.checkbox("Wi-Fi"): data['Wi-Fi Vendor'] = st.text_input("Вендор Wi-Fi:")
        if st.checkbox("QoS/Traffic Shaper"): data['QoS'] = "Да"

    if st.checkbox("Межсетевой экран (NGFW)"):
        ngfw_v = st.text_input("Вендор NGFW:")
        data['1.2.7. NGFW'] = f"Да ({ngfw_v})"
        score += 20

# 1.3 Серверы
st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
sc1, sc2 = st.columns(2)
with sc1:
    data['1.3.1. Физические серверы'] = st.number_input("Физические серверы:", min_value=0)
    data['Виртуализация'] = st.selectbox("Платформа:", ["VMware", "Hyper-V", "Proxmox", "KVM/Astra", "Нет"])
with sc2:
    data['1.3.2. Виртуальные серверы'] = st.number_input("Виртуальные серверы:", min_value=0)
    data['ОС Серверов'] = st.multiselect("ОС:", ["Windows Server", "Linux (RHEL/Ubuntu/Debian)", "Astra Linux"])

# 1.4 СХД (ALL-FLASH ВОССТАНОВЛЕНО)
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Наличие СХД", key="st_sw"):
    stc1, stc2 = st.columns(2)
    with stc1:
        st_a = st.selectbox("Тип массива:", ["All-Flash (NVMe/SSD)", "Hybrid", "HDD Only"])
        st_v = st.text_input("Вендор СХД:")
    with stc2:
        st_p = st.multiselect("Протоколы:", ["FC", "iSCSI", "NFS", "SMB/S3"])
        st_cap = st.number_input("Емкость (ТБ):", min_value=0)
    data['1.4. СХД'] = f"{st_v} | {st_a} ({st_cap} TB)"
else:
    data['1.4. СХД'] = "Нет"

if st.checkbox("Резервное копирование"):
    v_br = st.text_input("Производитель системы резервного копирования:")
    data["Резервное копирование"] = f"Да ({v_br})"
    score += 20

st.divider()

# --- БЛОК 2: ИБ ---
st.header("Блок 2: Информационная Безопасность")
ib_list = {
    "EPP (Антивирус)": 10, "DLP (Утечки)": 15, "PAM (Доступ)": 10,
    "SIEM (Мониторинг)": 20, "VM (Уязвимости)": 10, "EDR/XDR": 15, "MFA (2FA)": 15
}
ib_col1, ib_col2 = st.columns(2)
for i, (name, pts) in enumerate(ib_list.items()):
    with (ib_col1 if i < 4 else ib_col2):
        if st.checkbox(name):
            v_ib = st.text_input(f"Вендор {name}:", key=f"v_{name}")
            data[name] = f"Да ({v_ib})"
            score += pts
        else: data[name] = "Нет"

st.divider()

# --- БЛОК 4: РАЗРАБОТКА (ВОССТАНОВЛЕНО) ---
st.header("Блок 4: Разработка и DevOps")
if st.toggle("Внутренняя разработка", key="dev_sw"):
    dc1, dc2 = st.columns(2)
    with dc1:
        data['4.1. Разработчики'] = st.number_input("Кол-во (чел):", min_value=0)
        data['4.3. Репозиторий'] = st.selectbox("Git:", ["GitLab", "GitHub", "Bitbucket", "Нет"])
        if st.checkbox("Code Review"): data['Code Review'] = "Да"
    with dc2:
        data['4.4. CI/CD'] = st.selectbox("Автоматизация:", ["Jenkins", "GitLab CI", "GitHub Actions", "Нет"])
        data['4.5. Контейнеры'] = st.multiselect("Технологии:", ["Docker", "Kubernetes"])
        data['4.6. Окружения'] = st.multiselect("Среды:", ["Dev", "Stage", "Prod"])
else:
    data['4.1. Разработка'] = "Нет"

# --- ЭКСПОРТ (ЛОГИКА SIEM И РЕЗЕРВА) ---
def make_excel(c_info, results, total_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Report"
    
    blue_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    ws.merge_cells('A1:D2'); ws['A1'] = "ОТЧЕТ ПО ТЕХНИЧЕСКОМУ АУДИТУ 2026"; ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=14)

    row = 4
    for k, v in c_info.items():
        ws.cell(row=row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=row, column=2, value=str(v)); row += 1
    
    ws.cell(row=row, column=1, value="ЗРЕЛОСТЬ ИТ/ИБ:").font = Font(bold=True)
    ws.cell(row=row, column=2, value=f"{min(total_score, 100)}%"); row += 3

    for i, h in enumerate(["Параметр", "Значение", "Статус", "Рекомендация"], 1):
        cell = ws.cell(row=row, column=i, value=h); cell.fill = blue_fill; cell.font = white_font; cell.border = border
    
    row += 1
    n_arm = results.get('1.1. Всего АРМ', 0)
    n_srv = results.get('1.3.1. Физические серверы', 0) + results.get('1.3.2. Виртуальные серверы', 0)
    has_edr = results.get('EDR/XDR') != "Нет"

    for k, v in results.items():
        if "ОС" in k or "Виртуализация" in k: continue
        ws.cell(row=row, column=1, value=k).border = border
        ws.cell(row=row, column=2, value=str(v)).border = border
        
        status, rec, color = "В норме", "Ок.", "000000"
        if k == '1.2.2. Резервный канал' and v == "Нет":
            status, rec, color = "КРИТИЧНО", "Единая точка отказа. Требуется резервирование.", "FF0000"
        elif v == "Нет":
            if k == "SIEM (Мониторинг)" and n_arm < 100 and n_srv < 20 and not has_edr:
                status, rec, color = "ВНИМАНИЕ", "SIEM не критичен для текущего масштаба.", "FFC000"
            else:
                status, rec, color = "РИСК", "Рекомендуется внедрение для минимизации угроз.", "FF0000"

        st_c = ws.cell(row=row, column=3, value=status); st_c.font = Font(color=color, bold=True); st_c.border = border
        ws.cell(row=row, column=4, value=rec).border = border
        row += 1

    for c, w in {'A': 30, 'B': 30, 'C': 15, 'D': 50}.items(): ws.column_dimensions[c].width = w
    wb.save(output); return output.getvalue()

# --- ФИНАЛ ---
st.divider()
is_valid = all([client_info.get('Наименование компании'), client_info.get('Email'), client_info.get('Город')])

if st.button("📊 Сформировать экспертный отчет", disabled=not is_valid):
    report = make_excel(client_info, data, score)
    try:
        requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                      data={"chat_id": CHAT_ID, "caption": f"Аудит: {client_info['Наименование компании']}"}, 
                      files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report)})
    except: pass
    st.success("Отчет сформирован!")
    st.download_button("📥 Скачать Excel", report, f"Audit_{client_info['Наименование компании']}.xlsx")
elif not is_valid:
    st.warning("Заполните Город, Компанию и Email (обязательные поля).")
