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

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v6.7")

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
    
    # Исправленный блок Email
    custom_email = st.checkbox("Email отличается от домена сайта")
    if custom_email:
        client_info['Email'] = st.text_input("Email контактного лица:*", placeholder="info@other-domain.com")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        if clean_domain and "." in clean_domain:
            st.write("Email контактного лица:*")
            e_c1, e_c2 = st.columns([1, 2])
            with e_c1:
                prefix = st.text_input("Логин", placeholder="info", label_visibility="collapsed", key="em_pref")
            with e_c2:
                st.markdown(f"<div style='padding-top: 5px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
            client_info['Email'] = f"{prefix}@{clean_domain}" if prefix else ""
        else:
            client_info['Email'] = ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    
    # ВОЗВРАТ КОДА ТЕЛЕФОНА
    st.write("Контактный телефон:*")
    p_c1, p_c2 = st.columns([1, 2])
    codes = [("🇰🇿 +7", "+7"), ("🇷🇺 +7", "+7"), ("🇺🇿 +998", "+998")]
    selected_code = p_c1.selectbox("Код", codes, format_func=lambda x: x[0], label_visibility="collapsed")
    phone_val = p_c2.text_input("Номер", placeholder="707 000 00 00", label_visibility="collapsed")
    client_info['Телефон'] = f"{selected_code[1]} {phone_val}" if phone_val else ""

st.divider()

# --- БЛОК 1: ИТ ---
st.header("Блок 1: Информационные технологии")

# 1.1 АРМ
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)
data['1.1. Всего АРМ'] = total_arm

# 1.2 Сетевая инфраструктура
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сеть", key="net_t", value=True):
    net_types = ["Нет", "Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL"]
    col_n1, col_n2 = st.columns(2)
    with col_n1:
        data['1.2.1. Основной канал'] = st.selectbox("Основной канал:", net_types[1:])
    with col_n2:
        # Резервный канал по умолчанию "Нет"
        data['1.2.2. Резервный канал'] = st.selectbox("Резервный канал:", net_types, index=0)

    st.write("**Оборудование:**")
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.checkbox("Ядро сети (Core)"): data['Core'] = st.text_input("Производитель Core:")
    with c2:
        if st.checkbox("Коммутаторы L2/L3"): data['Switches'] = st.number_input("Кол-во:", min_value=0)
    with c3:
        if st.checkbox("Wi-Fi"): data['Wi-Fi Vendor'] = st.text_input("Производитель AP:")

    if st.checkbox("Межсетевой экран (NGFW)"):
        v_ng = st.text_input("Производитель NGFW (Fortigate/CheckPoint):")
        data['1.2.7. NGFW'] = f"Да ({v_ng})"
        score += 20

# 1.3 Серверы
st.write("---")
st.subheader("1.3. Серверы и Виртуализация")
data['1.3.1. Физические серверы'] = st.number_input("Физические серверы:", min_value=0)
data['1.3.2. Виртуальные серверы'] = st.number_input("Виртуальные серверы:", min_value=0)
data['Виртуализация'] = st.selectbox("Платформа:", ["VMware", "Hyper-V", "Proxmox", "KVM", "Нет"])

# 1.4 СХД (ВОССТАНОВЛЕНО)
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Наличие СХД", key="st_t"):
    s_c1, s_c2 = st.columns(2)
    with s_c1:
        st_arch = st.selectbox("Тип массива:", ["All-Flash (NVMe/SSD)", "Hybrid", "HDD Only"])
        st_vendor = st.text_input("Производитель СХД:")
    with s_c2:
        st_cap = st.number_input("Емкость (ТБ):", min_value=0)
        st_prot = st.multiselect("Протоколы:", ["FC", "iSCSI", "NFS", "SMB"])
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
    "EPP (Антивирус)": 10, "DLP (Утечки)": 15, "PAM (Доступ)": 10,
    "SIEM (Мониторинг)": 20, "VM (Уязвимости)": 10, "EDR/XDR": 15,
    "WAF (Защита Web)": 10, "MFA (2FA)": 15
}
col_ib1, col_ib2 = st.columns(2)
for i, (name, pts) in enumerate(ib_tools.items()):
    with (col_ib1 if i < 4 else col_ib2):
        if st.checkbox(name):
            v_i = st.text_input(f"Производитель {name}:", key=f"vi_{name}")
            data[name] = f"Да ({v_i})"
            score += pts
        else:
            data[name] = "Нет"

st.divider()

# --- БЛОК 4: РАЗРАБОТКА (ВОССТАНОВЛЕНО) ---
st.header("Блок 4: Разработка и DevOps")
if st.toggle("Внутренняя разработка", key="dev_t"):
    d1, d2 = st.columns(2)
    with d1:
        data['4.1. Разработчики'] = st.number_input("Штат (чел):", min_value=0)
        data['4.3. Репозиторий'] = st.selectbox("Git:", ["GitLab", "GitHub", "Bitbucket", "Нет"])
    with d2:
        data['4.4. CI/CD'] = st.selectbox("Автоматизация:", ["Jenkins", "GitLab CI", "GitHub Actions", "Нет"])
        data['4.5. Среды'] = st.multiselect("Наличие сред:", ["Dev", "Test", "Prod"])
else:
    data['4.1. Разработка'] = "Нет"

# --- ГЕНЕРАЦИЯ EXCEL ---
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
        c = ws.cell(row=row, column=i, value=h)
        c.fill = h_fill; c.font = w_font; c.border = brd
    
    row += 1
    n_arm = results.get('1.1. Всего АРМ', 0)
    n_srv = results.get('1.3.1. Физические серверы', 0) + results.get('1.3.2. Виртуальные серверы', 0)
    has_edr = results.get('EDR/XDR') != "Нет"

    for k, v in results.items():
        ws.cell(row=row, column=1, value=k).border = brd
        ws.cell(row=row, column=2, value=str(v)).border = brd
        
        status, rec, color = "В норме", "Ок.", "000000"

        if k == '1.2.2. Резервный канал' and v == "Нет":
            status, rec, color = "КРИТИЧНО", "Отсутствие резервного канала связи.", "FF0000"
        elif v == "Нет":
            if k == "SIEM (Мониторинг)":
                if n_arm < 100 and n_srv < 20 and not has_edr:
                    status, rec, color = "ВНИМАНИЕ", "Для малого масштаба SIEM не критичен.", "FFC000"
                else: status, rec, color = "КРИТИЧНО", "Необходим мониторинг инцидентов.", "FF0000"
            elif k == "1.4. СХД" and n_srv > 10:
                status, rec, color = "КРИТИЧНО", "Риск простоя из-за отсутствия СХД.", "FF0000"
            else: status, rec, color = "РИСК", "Рекомендуется внедрение.", "FF0000"

        st_c = ws.cell(row=row, column=3, value=status); st_c.font = Font(color=color, bold=True); st_c.border = brd
        ws.cell(row=row, column=4, value=rec).border = brd
        row += 1

    for c, w in {'A': 30, 'B': 30, 'C': 15, 'D': 50}.items(): ws.column_dimensions[c].width = w
    wb.save(output); return output.getvalue()

# --- ФИНАЛ ---
st.divider()
if st.button("📊 Сформировать экспертный отчет"):
    if not all([client_info['Наименование компании'], client_info['Email']]):
        st.error("Заполните обязательные поля (Компания, Email)!")
    else:
        report = generate_report(client_info, data, score)
        try:
            requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                          data={"chat_id": CHAT_ID, "caption": f"Аудит: {client_info['Наименование компании']}"}, 
                          files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report)})
        except: pass
        st.success("Отчет сформирован!")
        st.download_button("📥 Скачать Excel", report, f"Audit_{client_info['Наименование компании']}.xlsx")
