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

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v7.4 Gold")

# --- ИНСТРУКЦИЯ ДЛЯ ПОЛЬЗОВАТЕЛЯ ---
with st.expander("📖 Инструкция по заполнению (нажмите, чтобы развернуть)"):
    st.markdown("""
    ### Руководство по проведению экспресс-аудита
    1.  **Общая информация:** Укажите корректные контактные данные.
    2.  **Заполнение блоков:** Пройдите по разделам (ИТ, ИБ, Web, Разработка).
    3.  **Логический контроль:** Система проверяет соответствие количества ОС общему числу АРМ и серверов.
    4.  **Результат:** После заполнения нажмите «Сформировать экспертный отчет» для получения Excel-файла с анализом рисков.
    """)

data = {}
client_info = {}
validation_errors = []
score = 0

# --- ШАПКА: ИНФОРМАЦИЯ О КЛИЕНТЕ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город:*")
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    site_input = st.text_input("Сайт компании:*", placeholder="example.kz")
    client_info['Сайт компании'] = site_input

    custom_email_mode = st.checkbox("Email отличается от сайта")
    if custom_email_mode:
        client_info['Email'] = st.text_input("Email контактного лица:*")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица:*")
            e_col1, e_col2 = st.columns([2, 3])
            with e_col1:
                email_prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="email_pre")
            with e_col2:
                st.markdown(f"<div style='padding-top: 5px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
            client_info['Email'] = f"{email_prefix}@{clean_domain}" if email_prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    st.write("Контактный телефон:*")
    p_col1, p_col2 = st.columns([1, 2])
    country_codes = [("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"), ("🇰🇬 +996", "+996")]
    selected_code = p_col1.selectbox("Код", country_codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_num = p_col2.text_input("Номер", placeholder="777 777 77 77", label_visibility="collapsed")
    client_info['Телефон'] = f"{selected_code[1]} {phone_num}" if phone_num else ""

st.divider()

# --- БЛОК 1: ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ ---
st.header("Блок 1: Информационные технологии")

# 1.1 Конечные точки (АРМ) - Сопоставление количества и ОС
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

sum_arm = sum(arm_counts.values())
if total_arm > 0 and sum_arm != total_arm:
    st.warning(f"⚠️ Несоответствие: Сумма по ОС ({sum_arm}) не равна общему количеству ({total_arm})")
    validation_errors.append("Ошибка распределения АРМ по ОС")
for k, v in arm_counts.items(): data[f"АРМ: {k}"] = v

# 1.2 Сетевая инфраструктура
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сетевая инфраструктура", key="net_toggle", value=True):
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "Нет"]
    col_n1, col_n2 = st.columns(2)
    with col_n1:
        data['1.2.1. Основной канал'] = st.selectbox("Тип (основной):", net_types[:-1])
        data['Скорость осн.'] = st.number_input("Скорость (Mbit/s):", min_value=0, key="s1")
    with col_n2:
        data['1.2.2. Резервный канал'] = st.selectbox("Тип (резервный):", net_types, index=5)
        data['Скорость рез.'] = st.number_input("Скорость (Mbit/s):", min_value=0, key="s2")

    st.write("**Активное оборудование:**")
    c_net1, c_net2, c_net3 = st.columns(3)
    with c_net1:
        if st.checkbox("Ядро (Core Switch)"): data['Core'] = st.text_input("Вендор Core:")
        if st.checkbox("Маршрутизаторы"): data['Routers'] = st.number_input("Кол-во Routers:", min_value=0)
    with c_net2:
        if st.checkbox("Коммутаторы L3"): data['L3 Sw'] = st.number_input("Кол-во L3:", min_value=0)
        if st.checkbox("Коммутаторы L2"): data['L2 Sw'] = st.number_input("Кол-во L2:", min_value=0)
    with c_net3:
        if st.checkbox("Wi-Fi"): data['Wi-Fi'] = st.text_input("Вендор Wi-Fi:")
        if st.checkbox("QoS"): data['QoS'] = "Да"

    if st.checkbox("Межсетевой экран (NGFW)"):
        v_ng = st.text_input("Производитель NGFW:")
        data['1.2.7. NGFW'] = f"Да ({v_ng})"
        score += 20

# 1.3 Серверы (Детализация версий Windows отдельно)
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
    data['Linux (Server)'] = st.number_input("Linux (RHEL/CentOS/Ubuntu) (шт):", min_value=0)
    data['Astra Linux'] = st.number_input("Astra Linux (шт):", min_value=0)

if st.checkbox("Резервное копирование"):
    v_b = st.text_input("Вендор Backup:")
    data["Резервное копирование"] = f"Да ({v_b})"
    score += 20

# 1.4 СХД
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Есть собственная СХД", key="st_t"):
    stc1, stc2 = st.columns(2)
    with stc1:
        st_arch = st.selectbox("Тип:", ["All-Flash (NVMe/SSD)", "Hybrid", "HDD Only"])
        st_v = st.text_input("Вендор СХД:")
        if st.checkbox("HA (Дублирование контроллеров)"): data['СХД HA'] = "Да"
    with stc2:
        st_cap = st.number_input("Полезная емкость (ТБ):", min_value=0)
        data['1.4. СХД'] = f"{st_v} | {st_arch} ({st_cap} TB)"
else: data['1.4. СХД'] = "Нет"

# 1.5 ИС и WEB (Восстановлено: ERP, CRM, Billing, Web)
st.write("---")
st.subheader("1.5. Информационные системы и Web")
is1, is2 = st.columns(2)
with is1:
    data['ERP/CRM'] = st.text_input("ERP / CRM системы (1С, SAP и др.):")
    data['Billing'] = st.text_input("Биллинг / Фин. системы:")
with is2:
    data['Web External'] = st.text_input("Внешние сайты/порталы:")
    data['Web Internal'] = st.text_input("Внутренние ресурсы (Intranet):")

st.divider()

# --- БЛОК 2: ИБ (Восстановленные продукты ИБ) ---
st.header("Блок 2: Информационная Безопасность")
ib_tools = {
    "EPP (Антивирус)": 10, "DLP (Защита утечек)": 15, "PAM (Привилегии)": 10,
    "SIEM (Мониторинг)": 20, "VM (Уязвимости)": 10, "EDR/XDR": 15, 
    "WAF (Защита Web)": 10, "MFA (2FA)": 15
}
ib1, ib2 = st.columns(2)
for i, (name, pts) in enumerate(ib_tools.items()):
    with (ib1 if i < 4 else ib2):
        if st.checkbox(name):
            v_i = st.text_input(f"Вендор {name}:", key=f"ib_{name}")
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
        data['4.2. Стек'] = st.text_input("Стек:")
        data['4.3. Git'] = st.selectbox("Git:", ["GitLab", "GitHub", "Bitbucket", "Нет"])
    with d2:
        data['4.4. CI/CD'] = st.selectbox("CI/CD:", ["Jenkins", "GitLab CI", "Нет"])
        data['4.5. Контейнеры'] = st.multiselect("Среды:", ["Docker", "Kubernetes"])
        if st.checkbox("SAST/DAST"): data['SAST/DAST'] = "Да"
else: data['4.1. Разработка'] = "Нет"

# --- ГЕНЕРАЦИЯ EXCEL (Логика отчета) ---
def make_expert_excel(c_info, results, total_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit_Report"
    
    h_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    w_font = Font(color="FFFFFF", bold=True)
    brd = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    ws.merge_cells('A1:D2'); ws['A1'] = "ЭКСПЕРТНЫЙ ТЕХНИЧЕСКИЙ ОТЧЕТ 2026"; ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
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
        # Скрываем детальные ОС из сводной таблицы для чистоты
        if any(x in k for x in ["Win Srv", "Linux (", "Astra", "АРМ:"]): continue
        
        ws.cell(row=row, column=1, value=k).border = brd
        ws.cell(row=row, column=2, value=str(v)).border = brd
        
        status, rec, color = "В норме", "Ок.", "000000"
        if k == '1.2.2. Резервный канал' and v == "Нет":
            status, rec, color = "КРИТИЧНО", "Единая точка отказа связи.", "FF0000"
        elif v == "Нет":
            if k == "SIEM (Мониторинг)":
                if n_arm > 100 or n_srv > 20:
                    status, rec, color = "КРИТИЧНО", "Необходим мониторинг при вашем масштабе.", "FF0000"
                else: status, rec, color = "ВНИМАНИЕ", "SIEM рекомендован при росте сети.", "FFC000"
            else: status, rec, color = "РИСК", "Рекомендуется внедрение защиты.", "FF0000"

        st_c = ws.cell(row=row, column=3, value=status); st_c.font = Font(color=color, bold=True); st_c.border = brd
        ws.cell(row=row, column=4, value=rec).border = brd
        row += 1

    for c, w in {'A': 35, 'B': 30, 'C': 15, 'D': 55}.items(): ws.column_dimensions[c].width = w
    wb.save(output); return output.getvalue()

# --- ФИНАЛ ---
st.divider()
is_ready = all([client_info.get('Наименование компании'), client_info.get('Email'), client_info.get('Город'), client_info.get('Телефон')])

if validation_errors:
    st.error(f"🚨 Исправьте ошибки: {', '.join(validation_errors)}")

if st.button("📊 Сформировать экспертный отчет", disabled=not is_ready or len(validation_errors) > 0):
    with st.spinner("Генерация..."):
        report = make_expert_excel(client_info, data, score)
        try:
            requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                          data={"chat_id": CHAT_ID, "caption": f"Аудит: {client_info['Наименование компании']}"}, 
                          files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report)})
        except: pass
        st.success("Отчет готов!")
        st.download_button("📥 Скачать экспертный Excel", report, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Khalil Audit System v7.4 Gold | Almaty 2026")
