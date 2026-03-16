import streamlit as st
import pandas as pd
import os
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage

# --- 1. НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

# --- 2. ЛОГОТИП И КОНТАКТЫ ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade | IT Audit & Consulting")

st.markdown("### Мы поможем Вам стать лучше!**")
st.divider()

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026)")

data = {}
score = 0

# --- БЛОК 1: ОБЩАЯ ИНФОРМАЦИЯ ---
st.header("Блок 1: Общая информация")

st.subheader("1.1. Конечные точки (АРМ)")
total_arm = st.number_input("Общее количество АРМ (шт):", min_value=0, step=1, key="total_arm_val")
data['1.1. Всего АРМ'] = total_arm

selected_os_arm = st.multiselect("Выберите ОС на АРМ:", ["Windows", "Linux", "macOS", "Другое"], key="ms_arm_list")
if selected_os_arm:
    for os_item in selected_os_arm:
        count_arm = st.number_input(f"Количество АРМ на {os_item}:", min_value=0, step=1, key=f"arm_cnt_{os_item}")
        data[f"ОС АРМ ({os_item})"] = count_arm

st.write("---")
st.subheader("1.2. Серверы")
col_s1, col_s2 = st.columns(2)
with col_s1:
    phys_servers = st.number_input("Количество физических серверов:", min_value=0, step=1, key="phys_srv")
    data['1.2. Физические серверы'] = phys_servers
with col_s2:
    virt_servers = st.number_input("Количество виртуальных серверов:", min_value=0, step=1, key="virt_srv")
    data['1.2. Виртуальные серверы'] = virt_servers

selected_os_srv = st.multiselect("Выберите ОС серверов:", ["Windows Server", "Linux", "Unix", "Другое"], key="ms_srv_list")
if selected_os_srv:
    for os_s in selected_os_srv:
        count_srv = st.number_input(f"Количество серверов на {os_s}:", min_value=0, step=1, key=f"srv_cnt_{os_s}")
        data[f"ОС Сервера ({os_s})"] = count_srv

st.write("---")
col_v1, col_v2 = st.columns(2)
with col_v1:
    st.subheader("1.3. Виртуализация")
    data['1.3. Виртуализация'] = st.multiselect("Системы виртуализации:", ["VMware", "Hyper-V", "Proxmox", "KVM", "Другое", "Нет"], key="virt_sys")
with col_v2:
    st.subheader("1.4. Почтовая система")
    data['1.4. Почта'] = st.selectbox("Тип почты:", ["Exchange (On-Prem)", "Microsoft 365", "Google Workspace", "Yandex/Mail.ru Cloud", "Собственный сервер", "Нет"], key="mail_sys")

st.write("---")
st.subheader("1.5. Внутренние Информационные системы")
has_is = st.checkbox("Есть ли внутренние Информационные системы (1C, ERP, CRM)?", key="is_15_chk")
data['1.5. Внутренние ИС'] = st.text_input("Перечислите их через запятую:", key="is_input_field") if has_is else "Нет"

st.subheader("1.6. Система мониторинга")
has_mon = st.checkbox("Есть ли система мониторинга?", key="mon_16_chk")
if has_mon:
    data['1.6. Мониторинг'] = st.selectbox("Выберите систему:", ["Open-source решение", "Коммерческое ПО", "Облачный мониторинг", "Другое"], key="mon_select_field")
else:
    data['1.6. Мониторинг'] = "Нет"

# --- БЛОК 2: СЕТЬ ---
st.header("Блок 2: Сетевая инфраструктура и Интернет")
if st.toggle("Своя сетевая инфраструктура", key="net_block_toggle"):
    net_types = ["Оптика", "Радиорелейная", "Спутник", "4G/5G", "Starlink"]
    c_n1, c_n2 = st.columns(2)
    with c_n1:
        data['2.1. Основной канал'] = st.selectbox("Тип канала (Основной):", net_types, key="main_net_type")
        data['2.1. Скорость осн. (mbit/s)'] = st.number_input("Скорость основного канала (mbit/s):", min_value=0, key="main_net_speed")
    with c_n2:
        data['2.1. Резервный канал'] = st.selectbox("Тип канала (Резервный):", ["Нет"] + net_types, key="res_net_type")
        data['2.1. Скорость рез. (mbit/s)'] = st.number_input("Скорость резервного канала (mbit/s):", min_value=0, key="res_net_speed")

    data['2.4. NGFW'] = st.text_input("Вендор Межсетевого экрана (NGFW):", key="ngfw_vendor_input")
    if data['2.4. NGFW']: score += 20

    has_wifi = st.checkbox("Используется Wi-Fi?", key="wifi_usage_chk")
    if has_wifi:
        if st.checkbox("Есть ли Wi-Fi контроллер?", key="wifi_ctrl_exists"):
            data['2.5. Контроллер'] = st.text_input("Модель контроллера:", key="wifi_ctrl_model")
        else:
            data['2.5. Контроллер'] = "Без контроллера"
        data['2.5. Число точек'] = st.number_input("Кол-во точек доступа:", min_value=0, key="wifi_ap_count")

# --- БЛОК 3: ИБ ---
st.header("Блок 3: Информационная Безопасность")
if st.toggle("Есть отдел Информационной безопасности", key="ib_block_main_toggle"):
    ib_list = {"DLP (Защита от утечек)": 15, "PAM (Контроль доступа)": 10, "SIEM/SOC (Мониторинг ИБ)": 20, "WAF (Защита Web)": 10, "EDR/Antimalware": 15, "Резервное копирование": 20}
    for label, pts in ib_list.items():
        c_ib1, c_ib2 = st.columns([1, 2])
        is_on = c_ib1.checkbox(label, key=f"ib_v2_{label}")
        if is_on:
            v_name = c_ib2.text_input(f"Вендор {label}:", key=f"v_name_{label}")
            data[label], score = f"Да ({v_name if v_name else 'не указан'})", score + pts
        else:
            data[label] = "Нет"
    
    if st.checkbox("3.7. Другое", key="ib_other_toggle"):
        data['3.7. Прочие системы ИБ'] = st.text_area("Дополнительно:", key="ib_other_input")

# --- БЛОК 4 И 5 ---
st.header("Блок 4: Web-ресурсы")
if st.toggle("Есть свои Web-ресурсы", key="web_block_toggle"):
    data['4.1. Хостинг'] = st.selectbox("Где размещены сайты:", ["Собственный ЦОД", "Облако", "Нет сайтов"], key="web_hosting")
    data['4.2. CMS'] = st.text_input("Используемые CMS:", key="web_cms")

st.header("Блок 5: Разработка")
if st.toggle("Своя разработка", key="dev_block_toggle"):
    data['5.1. Разработчики'] = st.number_input("Кол-во разработчиков:", min_value=0, key="dev_count")
    data['5.3. CI/CD'] = st.checkbox("Используется ли CI/CD?", key="dev_cicd")

# --- ЭКСПЕРТНАЯ ЛОГИКА ---
def get_recommendation(key, value):
    db = {
        "Нет": "Критический риск. Отсутствие данной системы снижает прозрачность и безопасность ИТ-инфраструктуры.",
        "1.6. Мониторинг": "Рекомендуется внедрение систем мониторинга для проактивного обнаружения сбоев.",
        "2.1. Резервный канал": "Отсутствие резервирования канала — единая точка отказа. Рекомендуется дублирование связи.",
        "Резервное копирование": "КРИТИЧНО: Данные без бэкапа не существуют. Рекомендуется стратегия 3-2-1.",
        "WAF (Защита Web)": "Необходимо для защиты публичных ресурсов от атак прикладного уровня.",
        "PAM (Контроль доступа)": "Рекомендуется для аудита действий привилегированных пользователей.",
        "EDR/Antimalware": "Рекомендуются решения класса EDR для борьбы со сложными угрозами."
    }
    v_str = str(value)
    if "Нет" in v_str or v_str == "0":
        return db.get(key, db["Нет"])
    return "Конфигурация в норме. Рекомендуется плановое обслуживание."

# --- ЭКСЕЛЬ ГЕНЕРАЦИЯ ---
def make_excel(results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Анализ Аудита"

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    risk_fill = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    ws.merge_cells('A1:D1')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ: ТЕХНИЧЕСКИЙ АУДИТ 2026"
    ws['A1'].font = Font(bold=True, size=14, color="1F4E78")
    ws['A1'].alignment = Alignment(horizontal='center')

    ws['A3'] = "ИНДЕКС ЗРЕЛОСТИ:"
    ws['B3'] = f"{final_score} / 100"

    headers = ["Параметр", "Значение", "Анализ", "Рекомендация эксперта"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=5, column=i, value=h)
        cell.fill = header_fill
        cell.font = white_font

    row = 6
    for k, v in results.items():
        ws.cell(row=row, column=1, value=k).border = border
        ws.cell(row=row, column=2, value=str(v)).border = border
        
        rec = get_recommendation(k, v)
        status = "РИСК" if "Нет" in str(v) or "Рекомендуется" in rec else "ОК"
        
        sc = ws.cell(row=row, column=3, value=status)
        sc.border = border
        if status == "РИСК":
            sc.font = Font(color="FF0000", bold=True)
            ws.cell(row=row, column=1).fill = risk_fill

        rc = ws.cell(row=row, column=4, value=rec)
        rc.border = border
        rc.alignment = Alignment(wrap_text=True)
        row += 1

    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 50

    if os.path.exists("logo.png"):
        try:
            img = OpenpyxlImage("logo.png")
            img.height = 50; img.width = 150
            ws.add_image(img, 'D1')
        except: pass

    wb.save(output)
    return output.getvalue()

# --- ФИНАЛ ---
st.divider()
if st.button("📊 Сформировать экспертный отчет", key="final_btn_report"):
    if not data:
        st.error("Данные не заполнены!")
    else:
        f_score = min(score, 100)
        report = make_excel(data, f_score)
        st.success(f"Аналитика готова! Индекс: {f_score}/100")
        st.download_button("📥 Скачать экспертный Excel", report, "Audit_Expert_2026.xlsx", key="final_dl_btn")

st.info("Разработано Ivan Rudoy. По вопросам системной интеграции — звоните!")
