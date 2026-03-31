import streamlit as st
import pandas as pd
import os
import requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026 Gold", layout="wide", page_icon="🛡️")

# Telegram Secrets
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- 2. ВИЗУАЛЬНАЯ ЧАСТЬ ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### 🛡️ Технический аудит: Золотой образ + Логика v7.3")
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
    
    # Email: Логика сборки
    custom_email = st.checkbox("Email отличается от домена сайта")
    if custom_email:
        client_info['Email'] = st.text_input("Email контактного лица:*")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица:*")
            e_c1, e_c2 = st.columns([1, 2])
            with e_c1:
                prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="email_pref")
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

# 1.1 АРМ (ЗОЛОТОЙ СТАНДАРТ)
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)
data['1.1. Всего АРМ'] = total_arm

arm_counts = {}
ac1, ac2, ac3 = st.columns(3)
with ac1:
    arm_counts['Windows 10'] = st.number_input("Windows 10 (шт):", min_value=0)
    arm_counts['Windows 11'] = st.number_input("Windows 11 (шт):", min_value=0)
with ac2:
    arm_counts['Linux Desktop'] = st.number_input("Linux Desktop (шт):", min_value=0)
    arm_counts['macOS'] = st.number_input("macOS (шт):", min_value=0)
with ac3:
    arm_counts['Legacy (7/8/XP)'] = st.number_input("Legacy (7/8/XP) (шт):", min_value=0)

if sum(arm_counts.values()) != total_arm and total_arm > 0:
    st.warning(f"⚠️ Сумма по ОС ({sum(arm_counts.values())}) не совпадает с общим числом ({total_arm})")
for os_name, count in arm_counts.items():
    data[f"АРМ: {os_name}"] = count

# 1.2 СЕТИ (ПОЛНОЕ)
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
col_n1, col_n2 = st.columns(2)
net_types = ["Нет", "Оптика", "Радиорелейная", "Спутник", "4G/5G/Starlink", "ADSL"]
with col_n1:
    data['1.2.1. Основной канал'] = st.selectbox("Тип осн. канала:", net_types[1:])
    data['Скорость осн. (Mbit)'] = st.number_input("Скорость осн. канала:", min_value=0)
with col_n2:
    data['1.2.2. Резервный канал'] = st.selectbox("Тип рез. канала:", net_types, index=0)
    data['Скорость рез. (Mbit)'] = st.number_input("Скорость рез. канала:", min_value=0)

st.write("**Оборудование:**")
c_net1, c_net2, c_net3 = st.columns(3)
with c_net1:
    if st.checkbox("Ядро (Core Switch)"): data['Core'] = st.text_input("Вендор Core:")
    if st.checkbox("Маршрутизаторы"): data['Routers'] = st.number_input("Кол-во Routers:", min_value=0)
with c_net2:
    if st.checkbox("L3 Коммутаторы"): data['L3 Sw'] = st.number_input("Кол-во L3:", min_value=0)
    if st.checkbox("L2 Коммутаторы"): data['L2 Sw'] = st.number_input("Кол-во L2:", min_value=0)
with c_net3:
    if st.checkbox("Wi-Fi"): data['Wi-Fi'] = st.text_input("Вендор Wi-Fi:")
    if st.checkbox("QoS"): data['QoS'] = "Да"

if st.checkbox("Межсетевой экран (NGFW)"):
    v_ng = st.text_input("Производитель NGFW:")
    data['1.2.7. NGFW'] = f"Да ({v_ng})"
    score += 20

# 1.3 СЕРВЕРЫ (ДЕТАЛИЗИРОВАННЫЕ ВЕРСИИ)
st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
sc1, sc2 = st.columns(2)
with sc1:
    data['1.3.1. Физические серверы'] = st.number_input("Физические серверы (шт):", min_value=0)
    data['Виртуализация'] = st.selectbox("Платформа:", ["VMware", "Hyper-V", "Proxmox", "KVM/Astra", "Нет"])
with sc2:
    data['1.3.2. Виртуальные серверы'] = st.number_input("Виртуальные серверы (шт):", min_value=0)

st.write("**Детализация ОС Серверов:**")
srv_os1, srv_os2 = st.columns(2)
with srv_os1:
    data['Win Srv 2012/R2'] = st.number_input("Windows Server 2012/R2 (шт):", min_value=0)
    data['Win Srv 2016'] = st.number_input("Windows Server 2016 (шт):", min_value=0)
    data['Win Srv 2019'] = st.number_input("Windows Server 2019 (шт):", min_value=0)
    data['Win Srv 2022'] = st.number_input("Windows Server 2022 (шт):", min_value=0)
with srv_os2:
    data['Linux (RHEL/CentOS)'] = st.number_input("Linux (RHEL/CentOS) (шт):", min_value=0)
    data['Linux (Deb/Ubu)'] = st.number_input("Linux (Debian/Ubuntu) (шт):", min_value=0)
    data['Astra Linux'] = st.number_input("Astra Linux (шт):", min_value=0)

# 1.4 СХД (ЗОЛОТОЙ СТАНДАРТ)
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Наличие СХД", key="st_sw"):
    stc1, stc2 = st.columns(2)
    with stc1:
        st_arch = st.selectbox("Тип массива:", ["All-Flash (NVMe/SSD)", "Hybrid", "HDD Only"])
        st_vendor = st.text_input("Вендор СХД:")
        if st.checkbox("HA (Высокая доступность / 2 контроллера)"): data['СХД HA'] = "Да"
    with stc2:
        st_cap = st.number_input("Полезная емкость (ТБ):", min_value=0)
        data['1.4. СХД'] = f"{st_vendor} | {st_arch} ({st_cap} TB)"
else:
    data['1.4. СХД'] = "Нет"

if st.checkbox("Резервное копирование (Backup)"):
    v_back = st.text_input("Вендор Backup:")
    data["Резервное копирование"] = f"Да ({v_back})"
    score += 20

# 1.5 ИС и WEB (ВОССТАНОВЛЕНО)
st.write("---")
st.subheader("1.5. Информационные системы и Web-ресурсы")
is1, is2 = st.columns(2)
with is1:
    data['ERP/CRM'] = st.text_input("ERP / CRM системы (1С, SAP и др.):")
    data['Billing'] = st.text_input("Биллинг / Фин. системы:")
with is2:
    data['Web Ext'] = st.text_input("Внешние Web-ресурсы (Сайты/Порталы):")
    data['Web Int'] = st.text_input("Внутренние ресурсы (Intranet/Wiki):")

st.divider()

# --- БЛОК 2: ИБ (ЗОЛОТОЙ СПИСОК) ---
st.header("Блок 2: Информационная Безопасность")
ib_tools = {
    "EPP (Антивирус)": 10, "DLP (Утечки)": 15, "PAM (Доступ)": 10,
    "SIEM (Мониторинг)": 20, "VM (Уязвимости)": 10, "EDR/XDR": 15, 
    "WAF (Защита Web)": 10, "MFA (2FA)": 15
}
ib_c1, ib_c2 = st.columns(2)
for i, (name, pts) in enumerate(ib_tools.items()):
    with (ib_c1 if i < 4 else ib_c2):
        if st.checkbox(name):
            v_ib = st.text_input(f"Производитель {name}:", key=f"ib_{name}")
            data[name] = f"Да ({v_ib})"
            score += pts
        else:
            data[name] = "Нет"

st.divider()

# --- БЛОК 4: РАЗРАБОТКА ---
st.header("Блок 4: Разработка и DevOps")
if st.toggle("Внутренняя разработка", key="dev_sw"):
    dc1, dc2 = st.columns(2)
    with dc1:
        data['4.1. Разработчики'] = st.number_input("Кол-во разработчиков:", min_value=0)
        data['4.2. Стек'] = st.text_input("Основной стек:")
        data['4.3. Репозиторий'] = st.selectbox("Хранение кода:", ["GitLab", "GitHub", "Bitbucket", "Нет"])
    with dc2:
        data['4.4. CI/CD'] = st.selectbox("CI/CD:", ["Jenkins", "GitLab CI", "GitHub Actions", "Нет"])
        data['4.5. Контейнеры'] = st.multiselect("Технологии:", ["Docker", "Kubernetes"])
        data['4.6. Среды'] = st.multiselect("Окружения:", ["Dev", "Stage", "Prod"])
        if st.checkbox("Анализ кода (SAST/DAST)"): data['SAST/DAST'] = "Да"
else:
    data['4.1. Разработка'] = "Нет"

# --- ЛОГИКА ОТЧЕТА (EXCEL + АНАЛИТИКА) ---
def build_report(c_info, results, total_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit"
    
    blue_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    ws.merge_cells('A1:D2'); ws['A1'] = "ЭКСПЕРТНЫЙ ТЕХНИЧЕСКИЙ АУДИТ 2026"; ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=14)

    # Инфо клиента
    curr_row = 4
    for k, v in c_info.items():
        ws.cell(row=curr_row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=curr_row, column=2, value=str(v))
        curr_row += 1
    
    ws.cell(row=curr_row, column=1, value="ЗРЕЛОСТЬ ИНФРАСТРУКТУРЫ:").font = Font(bold=True)
    ws.cell(row=curr_row, column=2, value=f"{min(total_score, 100)}%"); curr_row += 3

    # Шапка таблицы
    headers = ["Параметр", "Значение", "Статус", "Рекомендация / Риск"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=curr_row, column=i, value=h)
        cell.fill = blue_fill; cell.font = white_font; cell.border = border
    
    curr_row += 1
    
    # Расчетные данные для алертов
    n_arm = results.get('1.1. Всего АРМ', 0)
    n_srv = results.get('1.3.1. Физические серверы', 0) + results.get('1.3.2. Виртуальные серверы', 0)
    has_edr = results.get('EDR/XDR') != "Нет"

    # Основной цикл по данным
    for key, val in results.items():
        # Пропускаем детальные версии ОС из общей таблицы для лаконичности (они в шапке)
        if any(x in key for x in ["Win Srv", "Linux (", "Astra", "АРМ:"]): continue
        
        ws.cell(row=curr_row, column=1, value=key).border = border
        ws.cell(row=curr_row, column=2, value=str(val)).border = border
        
        status, rec, color = "В норме", "Ок.", "000000"
        
        # ЛОГИКА АЛЕРТОВ
        if key == '1.2.2. Резервный канал' and val == "Нет":
            status, rec, color = "КРИТИЧНО", "Единая точка отказа. Требуется резервный канал связи.", "FF0000"
        
        elif val == "Нет":
            if key == "SIEM (Мониторинг)":
                if n_arm < 100 and n_srv < 20 and not has_edr:
                    status, rec, color = "ВНИМАНИЕ", "SIEM желателен, но не критичен при текущем масштабе.", "FFC000"
                else:
                    status, rec, color = "КРИТИЧНО", "Необходим мониторинг при текущем количестве узлов.", "FF0000"
            elif "ОС" not in key and "Разработка" not in key:
                status, rec, color = "РИСК", "Рекомендуется внедрение для снижения угроз.", "FF0000"

        st_cell = ws.cell(row=curr_row, column=3, value=status)
        st_cell.font = Font(color=color, bold=True); st_cell.border = border
        ws.cell(row=curr_row, column=4, value=rec).border = border
        curr_row += 1

    for c, w in {'A': 35, 'B': 30, 'C': 15, 'D': 55}.items(): ws.column_dimensions[c].width = w
    wb.save(output); return output.getvalue()

# --- ФИНАЛЬНАЯ КНОПКА ---
st.divider()
is_valid = all([client_info.get('Наименование компании'), client_info.get('Email'), client_info.get('Город'), client_info.get('Телефон')])

if st.button("📊 Сформировать экспертный отчет", disabled=not is_valid):
    report_file = build_report(client_info, data, score)
    try:
        requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                      data={"chat_id": CHAT_ID, "caption": f"Аудит: {client_info['Наименование компании']}"}, 
                      files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report_file)})
    except: pass
    st.success("Отчет успешно сформирован и отправлен!")
    st.download_button("📥 Скачать экспертный Excel", report_file, f"Audit_{client_info['Наименование компании']}.xlsx")
elif not is_valid:
    st.warning("⚠️ Для генерации отчета необходимо заполнить обязательные поля (Компания, Город, Email, Телефон).")
