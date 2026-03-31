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

# --- 2. ЛОГОТИП И КОНТАКТЫ ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### Мы поможем Вам стать лучше!")
st.divider()

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v6.8")

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
    
    # Email ОБЯЗАТЕЛЬНОЕ ПОЛЕ
    custom_email = st.checkbox("Email отличается от домена сайта")
    if custom_email:
        client_info['Email'] = st.text_input("Email контактного лица:*", placeholder="user@company.com")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица (обязательно):*")
            e_c1, e_c2 = st.columns([1, 2])
            with e_c1:
                prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="email_user")
            with e_c2:
                st.markdown(f"<div style='padding-top: 5px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
            client_info['Email'] = f"{prefix}@{clean_domain}" if prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    
    # РАСШИРЕННЫЕ КОДЫ ТЕЛЕФОНОВ
    st.write("Контактный телефон:*")
    p_c1, p_c2 = st.columns([1, 2])
    codes = [
        ("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998"), 
        ("🇰🇬 +996", "+996"), ("🇹🇯 +992", "+992"), ("🇹🇲 +993", "+993")
    ]
    selected_code = p_c1.selectbox("Код", codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_val = p_c2.text_input("Номер", placeholder="707 000 00 00", label_visibility="collapsed")
    client_info['Телефон'] = f"{selected_code[1]} {phone_val}" if phone_val else ""

st.divider()

# --- БЛОК 1: ИТ ---
st.header("Блок 1: Информационные технологии")

# 1.2 Сетевая инфраструктура
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сеть", key="net_t", value=True):
    net_types = ["Нет", "Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL"]
    col_n1, col_n2 = st.columns(2)
    with col_n1:
        data['1.2.1. Основной канал'] = st.selectbox("Основной канал:", net_types[1:])
    with col_n2:
        # Резервный канал по умолчанию "Нет"
        data['1.2.2. Резервный канал'] = st.selectbox("Резервный канал:", net_types, index=0)

    st.write("**Дополнительно по сети:**")
    nc1, nc2 = st.columns(2)
    with nc1:
        if st.checkbox("Управление трафиком (QoS/Bandwidth Management)"): data['QoS'] = "Да"
    with nc2:
        if st.checkbox("Межсетевой экран (NGFW)"):
            v_ng = st.text_input("Производитель NGFW:")
            data['1.2.7. NGFW'] = f"Да ({v_ng})"
            score += 20

# 1.4 СХД (ALL-FLASH И Т.Д.)
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Наличие СХД", key="st_t"):
    s_c1, s_c2 = st.columns(2)
    with s_c1:
        st_arch = st.selectbox("Тип массива:", ["All-Flash (NVMe/SSD)", "Hybrid (SSD+HDD)", "HDD Only"])
        st_vendor = st.text_input("Производитель СХД (Dell, HP, Huawei...):")
    with s_c2:
        st_cap = st.number_input("Полезная емкость (ТБ):", min_value=0)
        st_prot = st.multiselect("Протоколы:", ["FC", "iSCSI", "NFS", "SMB/S3"])
    data['1.4. СХД'] = f"{st_vendor} | {st_arch} ({st_cap} TB)"
else:
    data['1.4. СХД'] = "Нет"

if st.checkbox("Резервное копирование"):
    v_back = st.text_input("Производитель системы резервного копирования:")
    data["Резервное копирование"] = f"Да ({v_back})"
    score += 20

st.divider()

# --- БЛОК 2: ИБ ---
st.header("Блок 2: Информационная Безопасность")
ib_tools = {
    "EPP (Антивирус)": 10, "DLP (Утечки)": 15, "PAM (Привилегии)": 10,
    "SIEM (Мониторинг)": 20, "VM (Уязвимости)": 10, "EDR/XDR": 15, "MFA (2FA)": 15
}
col_ib1, col_ib2 = st.columns(2)
for i, (name, pts) in enumerate(ib_tools.items()):
    with (col_ib1 if i < 4 else col_ib2):
        if st.checkbox(name):
            v_i = st.text_input(f"Производитель {name}:", key=f"vi_{name}")
            data[name] = f"Да ({v_i})"
            score += pts
        else: data[name] = "Нет"

st.divider()

# --- БЛОК 4: РАЗРАБОТКА ---
st.header("Блок 4: Разработка и DevOps")
if st.toggle("Внутренняя разработка", key="dev_t"):
    d1, d2 = st.columns(2)
    with d1:
        data['4.1. Разработчики'] = st.number_input("Количество (чел):", min_value=0)
        data['4.3. Репозиторий'] = st.selectbox("Хранение кода:", ["GitLab", "GitHub", "Bitbucket", "Нет"])
        if st.checkbox("Code Review процесс"): data['Code Review'] = "Да"
    with d2:
        data['4.4. CI/CD'] = st.selectbox("Автоматизация:", ["Jenkins", "GitLab CI", "GitHub Actions", "Нет"])
        data['4.5. Контейнеризация'] = st.multiselect("Технологии:", ["Docker", "Kubernetes"])
        data['4.6. Среды'] = st.multiselect("Окружения:", ["Development", "Staging", "Production"])
else:
    data['4.1. Разработка'] = "Нет"

# --- ЛОГИКА ГЕНЕРАЦИИ (БЕЗ ИЗМЕНЕНИЙ) ---
def generate_report(c_info, results, total_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit"
    
    h_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    w_font = Font(color="FFFFFF", bold=True)
    brd = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    ws.merge_cells('A1:D2'); ws['A1'] = "ЭКСПЕРТНЫЙ ТЕХНИЧЕСКИЙ АУДИТ 2026"; ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
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
    # Логика алертов
    for k, v in results.items():
        ws.cell(row=row, column=1, value=k).border = brd
        ws.cell(row=row, column=2, value=str(v)).border = brd
        
        status, rec, color = "В норме", "Ок.", "000000"
        if k == '1.2.2. Резервный канал' and v == "Нет":
            status, rec, color = "КРИТИЧНО", "Высокий риск простоя. Рекомендуется резервный канал.", "FF0000"
        elif v == "Нет" and "ОС" not in k:
            status, rec, color = "РИСК", "Рекомендуется внедрение для минимизации угроз.", "FF0000"

        st_c = ws.cell(row=row, column=3, value=status); st_c.font = Font(color=color, bold=True); st_c.border = brd
        ws.cell(row=row, column=4, value=rec).border = brd
        row += 1

    for c, w in {'A': 30, 'B': 30, 'C': 15, 'D': 50}.items(): ws.column_dimensions[c].width = w
    wb.save(output); return output.getvalue()

# --- ФИНАЛ ---
st.divider()
# КНОПКА ЗАБЛОКИРОВАНА ПРИ ОТСУТСТВИИ EMAIL
is_ready = all([client_info.get('Наименование компании'), client_info.get('Email'), client_info.get('Город')])

if st.button("📊 Сформировать экспертный отчет", disabled=not is_ready):
    report = generate_report(client_info, data, score)
    try:
        requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                      data={"chat_id": CHAT_ID, "caption": f"Аудит: {client_info['Наименование компании']}"}, 
                      files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report)})
    except: pass
    st.success("Отчет сформирован!")
    st.download_button("📥 Скачать Excel", report, f"Audit_{client_info['Наименование компании']}.xlsx")
elif not is_ready:
    st.warning("⚠️ Для активации кнопки заполните обязательные поля: Город, Компания и Email.")
