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

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026) v6.4")

data = {}
client_info = {}
validation_errors = []
score = 0

# --- ШАПКА: ОБЩАЯ ИНФОРМАЦИЯ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город:*")
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    site_input = st.text_input("Сайт компании:*", placeholder="example.kz")
    client_info['Сайт компании'] = site_input
    
    if st.checkbox("Email отличается от домена сайта"):
        client_info['Email'] = st.text_input("Email контактного лица:*")
    else:
        clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]
        prefix = st.text_input("Логин (email):", placeholder="info")
        client_info['Email'] = f"{prefix}@{clean_domain}" if prefix and clean_domain else ""

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    client_info['Телефон'] = st.text_input("Контактный телефон:*", placeholder="+7 707 000 00 00")

st.divider()

# --- БЛОК 1: ИТ ---
st.header("Блок 1: Информационные технологии")

# 1.1 АРМ
st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1)
data['1.1. Всего АРМ'] = total_arm

# 1.2 Сеть (ИСПРАВЛЕНО: Резервный канал в "Нет" по дефолту)
st.write("---")
st.subheader("1.2. Сетевая инфраструктура")
if st.toggle("Своя сеть", key="net_t", value=True):
    net_types = ["Нет", "Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink", "ADSL"]
    col_n1, col_n2 = st.columns(2)
    with col_n1:
        data['1.2.1. Основной канал'] = st.selectbox("Основной канал:", net_types[1:])
    with col_n2:
        # По умолчанию "Нет"
        data['1.2.2. Резервный канал'] = st.selectbox("Резервный канал:", net_types, index=0)

    st.write("**Оборудование:**")
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.checkbox("Ядро сети (Core)"): data['Core'] = st.text_input("Производитель Core:")
    with c2:
        if st.checkbox("Коммутаторы L2/L3"): data['Switches'] = st.number_input("Кол-во:", min_value=0)
    with c3:
        if st.checkbox("Wi-Fi"): data['Wi-Fi'] = st.text_input("Производитель AP:")

    if st.checkbox("Межсетевой экран (NGFW)"):
        v_ng = st.text_input("Производитель NGFW:")
        data['1.2.7. NGFW'] = f"Да ({v_ng})"
        score += 20

# 1.3 Серверы
st.write("---")
st.subheader("1.3. Серверы")
data['1.3.1. Физические серверы'] = st.number_input("Физические серверы:", min_value=0)
data['1.3.2. Виртуальные серверы'] = st.number_input("Виртуальные серверы:", min_value=0)

# 1.4 СХД (ПОЛНОЕ ВОССТАНОВЛЕНИЕ)
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")
if st.toggle("Наличие СХД", key="storage_toggle"):
    s_col1, s_col2 = st.columns(2)
    with s_col1:
        st_arch = st.selectbox("Архитектура:", ["All-Flash (NVMe/SSD)", "Hybrid", "HDD Only"])
        st_vendor = st.text_input("Производитель СХД:")
    with s_col2:
        st_cap = st.number_input("Емкость (ТБ):", min_value=0)
        st_prot = st.multiselect("Протоколы:", ["FC", "iSCSI", "NFS", "SMB", "S3"])
    data['1.4. СХД'] = f"{st_vendor} | {st_arch} ({st_cap} TB)"
else:
    data['1.4. СХД'] = "Нет"

if st.checkbox("Резервное копирование"):
    v_b = st.text_input("Производитель системы резервного копирования:")
    data["Резервное копирование"] = f"Да ({v_b})"
    score += 20

st.divider()

# --- БЛОК 2: ИБ ---
st.header("Блок 2: Информационная Безопасность")
ib_tools = {
    "EPP (Антивирус)": 10, "DLP (Утечки)": 15, "PAM (Привилегии)": 10,
    "SIEM (Мониторинг)": 20, "VM (Уязвимости)": 10, "EDR/XDR": 15,
    "WAF (Защита Web)": 10, "MFA (2FA)": 15
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

# --- БЛОК 4: РАЗРАБОТКА (РАСШИРЕНО) ---
st.header("Блок 4: Разработка и DevOps")
if st.toggle("Внутренняя разработка", key="dev_toggle"):
    d1, d2 = st.columns(2)
    with d1:
        data['4.1. Кол-во разработчиков'] = st.number_input("Штат (чел):", min_value=0)
        data['4.2. Стек'] = st.text_input("Языки/Фреймворки:")
        data['4.3. Хранение кода'] = st.selectbox("Git-сервис:", ["GitLab", "GitHub", "Bitbucket", "Локальный Git", "Нет"])
    with d2:
        data['4.4. CI/CD'] = st.selectbox("Автоматизация сборок:", ["Jenkins", "GitLab CI", "GitHub Actions", "Нет"])
        data['4.5. Контейнеризация'] = st.multiselect("Технологии:", ["Docker", "Kubernetes", "Docker Swarm"])
        data['4.6. Среды'] = st.multiselect("Наличие окружений:", ["Dev", "Stage/Test", "Prod (Изолирован)"])
    
    if st.checkbox("Документация кода (Swagger/Wiki)"): data['4.7. Документация'] = "Да"
else:
    data['4.1. Разработка'] = "Нет"

# --- ГЕНЕРАЦИЯ EXCEL (НОВАЯ ЛОГИКА) ---
def generate_report(c_info, results, total_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit 2026"
    
    # Стили
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Шапка
    ws.merge_cells('A1:D2'); ws['A1'] = "ЭКСПЕРТНЫЙ ТЕХНИЧЕСКИЙ АУДИТ (2026)"; ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=14)

    row = 4
    for k, v in c_info.items():
        ws.cell(row=row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=row, column=2, value=str(v)); row += 1
    
    ws.cell(row=row, column=1, value="ЗРЕЛОСТЬ ИТ/ИБ:").font = Font(bold=True)
    ws.cell(row=row, column=2, value=f"{min(total_score, 100)}%"); row += 3

    cols = ["Параметр", "Значение", "Статус", "Рекомендация / Обоснование"]
    for i, name in enumerate(cols, 1):
        cell = ws.cell(row=row, column=i, value=name)
        cell.fill = header_fill; cell.font = white_font; cell.border = border
    
    row += 1
    
    n_arm = results.get('1.1. Всего АРМ', 0)
    n_srv = results.get('1.3.1. Физические серверы', 0) + results.get('1.3.2. Виртуальные серверы', 0)
    has_edr = results.get('EDR/XDR') != "Нет"

    for k, v in results.items():
        if "ОС" in k or "speed" in k: continue
        ws.cell(row=row, column=1, value=k).border = border
        ws.cell(row=row, column=2, value=str(v)).border = border
        
        status, rec, color = "В норме", "Риски под контролем.", "000000"

        # ЛОГИКА ПРОВЕРОК
        if k == '1.2.2. Резервный канал' and v == "Нет":
            status, rec, color = "КРИТИЧНО", "Единая точка отказа. Требуется резервный канал связи.", "FF0000"
        
        elif v == "Нет":
            if k == "SIEM (Мониторинг)":
                if n_arm < 100 and n_srv < 20 and not has_edr:
                    status, rec, color = "ВНИМАНИЕ", "Малый масштаб. Рекомендуется к приобретению при росте инфраструктуры.", "FFC000"
                else:
                    status, rec, color = "КРИТИЧНО", "Необходим централизованный контроль инцидентов.", "FF0000"
            elif k == "1.4. СХД":
                if n_srv > 10: status, rec, color = "КРИТИЧНО", "Без СХД невозможна высокая доступность сервисов.", "FF0000"
                else: status, rec, color = "ВНИМАНИЕ", "Рекомендуется к приобретению для надежности хранения.", "FFC000"
            elif k == "MFA (2FA)":
                status, rec, color = "КРИТИЧНО", "Рекомендуется к приобретению для защиты доступа.", "FF0000"
            else:
                status, rec, color = "РИСК", "Рекомендуется к приобретению для усиления ИБ.", "FF0000"

        st_cell = ws.cell(row=row, column=3, value=status)
        st_cell.font = Font(color=color, bold=True); st_cell.border = border
        ws.cell(row=row, column=4, value=rec).border = border
        row += 1

    for c, w in {'A': 30, 'B': 25, 'C': 15, 'D': 55}.items(): ws.column_dimensions[c].width = w
    wb.save(output); return output.getvalue()

# --- ВЫВОД ---
st.divider()
if st.button("📊 Сформировать экспертный отчет"):
    if not client_info['Наименование компании'] or not client_info['Email']:
        st.error("Заполните обязательные поля!")
    else:
        with st.spinner("Анализ..."):
            report = generate_report(client_info, data, score)
            # Отправка в ТГ
            try:
                requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                              data={"chat_id": CHAT_ID, "caption": f"Аудит: {client_info['Наименование компании']}"}, 
                              files={'document': (f"Audit_{client_info['Наименование компании']}.xlsx", report)})
            except: pass
            
            st.success("Отчет готов!")
            st.download_button("📥 Скачать Excel", report, f"Audit_{client_info['Наименование компании']}.xlsx")
