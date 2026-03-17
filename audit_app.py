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

st.markdown("### Мы поможем Вам стать лучше!**")
st.divider()

st.title("📋 Опросник: Технический аудит ИТ и ИБ (2026)")

data = {}
client_info = {}
score = 0

# --- НОВАЯ ШАПКА: ИНФОРМАЦИЯ О КЛИЕНТЕ ---
st.header("📍 Общая информация")
col_h1, col_h2 = st.columns(2)

with col_h1:
    client_info['Город'] = st.text_input("Город:*")
    # Используем дату из календаря, но в отчет пойдет строкой
    filling_date = st.date_input("Дата заполнения:*", value=datetime.now())
    client_info['Дата заполнения'] = filling_date.strftime("%d.%m.%Y")
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    client_info['Сайт компании'] = st.text_input("Сайт компании (необязательно):")

with col_h2:
    client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица:*")
    client_info['Должность'] = st.text_input("Должность:*")
    client_info['Контактный телефон'] = st.text_input("Контактный телефон:*")

st.divider()

# --- БЛОК 1: ТЕХНИЧЕСКАЯ ИНФРАСТРУКТУРА ---
st.header("Блок 1: Техническая часть")

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
if has_is:
    data['1.5. Внутренние ИС'] = st.text_input("Перечислите их через запятую:", key="is_input_field")
else:
    data['1.5. Внутренние ИС'] = "Нет"

st.subheader("1.6. Система мониторинга")
has_mon = st.checkbox("Есть ли система мониторинга?", key="mon_16_chk")
if has_mon:
    data['1.6. Мониторинг'] = st.selectbox("Выберите систему:", ["Zabbix", "Nagios", "PRTG", "Prometheus", "Другое"], key="mon_select_field")
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
    st.info("Отметьте внедренные системы и укажите их вендора:")
    ib_list = {
        "DLP (Защита от утечек)": 15,
        "PAM (Контроль доступа)": 10,
        "SIEM/SOC (Мониторинг ИБ)": 20,
        "WAF (Защита Web)": 10,
        "EDR/Antimalware": 15,
        "Резервное копирование": 20
    }
    for label, pts in ib_list.items():
        c_ib1, c_ib2 = st.columns([1, 2])
        with c_ib1:
            is_on = st.checkbox(label, key=f"ib_v2_{label}")
        if is_on:
            with c_ib2:
                v_name = st.text_input(f"Вендор {label}:", key=f"v_name_{label}")
                data[label] = f"Да ({v_name if v_name else 'не указан'})"
                score += pts
        else:
            data[label] = "Нет"
    
    st.write("---")
    if st.checkbox("3.7. Другое (дополнительные системы защиты)", key="ib_other_toggle"):
        data['3.7. Прочие системы ИБ'] = st.text_area("Перечислите все, что мы не учли:", key="ib_other_input")
    else:
        data['3.7. Прочие системы ИБ'] = "Нет"


# --- БЛОК 4: WEB-РЕСУРСЫ ---
st.header("Блок 4: Web-ресурсы")
if st.toggle("Есть свои Web-ресурсы", key="web_block_toggle"):
    data['4.1. Хостинг'] = st.selectbox("Где размещены сайты:", ["Собственный ЦОД", "Облако (Казахстан)", "Облако (Мировое)", "Нет сайтов"], key="web_hosting")
    data['4.2. CMS'] = st.text_input("Используемые CMS (Bitrix, WP и т.д.):", key="web_cms")
    data['4.3. СУБД'] = st.multiselect("Используемые базы данных:", ["PostgreSQL", "MySQL", "MS SQL", "Oracle", "MongoDB"], key="web_db")


# --- БЛОК 5: РАЗРАБОТКА ---
st.header("Блок 5: Разработка")
if st.toggle("Своя разработка", key="dev_block_toggle"):
    data['5.1. Разработчики'] = st.number_input("Количество разработчиков в штате:", min_value=0, key="dev_count")
    data['5.2. Стек'] = st.text_input("Основной стек технологий:", key="dev_stack")
    data['5.3. CI/CD'] = st.checkbox("Используется ли CI/CD?", key="dev_cicd")
    data['5.4. Контейнеры'] = st.text_input("Технологии (Docker/K8s):", key="dev_cont")


# --- ЭКСЕЛЬ ГЕНЕРАЦИЯ ---
def make_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Анализ Аудита"

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    ws.merge_cells('A1:C1')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ: ТЕХНИЧЕСКИЙ АУДИТ 2026"
    ws['A1'].alignment = Alignment(horizontal='center')
    ws['A1'].font = Font(bold=True, size=12)

    # Информация о компании
    current_row = 3
    for k, v in c_info.items():
        ws.cell(row=current_row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=current_row, column=2, value=str(v))
        current_row += 1

    current_row += 1
    ws.cell(row=current_row, column=1, value="ИНДЕКС ТЕХНИЧЕСКОЙ ЗРЕЛОСТИ:").font = Font(bold=True)
    ws.cell(row=current_row, column=2, value=f"{final_score} / 100")
    bg_color = "00B050" if final_score > 70 else "FFCC00" if final_score > 40 else "FF0000"
    ws.cell(row=current_row, column=2).fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    
    current_row += 2
    headers = ["Параметр", "Значение", "Анализ"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=current_row, column=i, value=h)
        cell.fill = header_fill
        cell.font = white_font

    current_row += 1
    for k, v in results.items():
        ws.cell(row=current_row, column=1, value=k).border = border
        ws.cell(row=current_row, column=2, value=str(v)).border = border
        
        status = "В норме"
        if "Нет" in str(v) or v == 0:
            status = "РИСК / ТРЕБУЕТ ВНИМАНИЯ"
            ws.cell(row=current_row, column=3).font = Font(color="FF0000", bold=True)
        ws.cell(row=current_row, column=3, value=status).border = border
        current_row += 1

    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 25
    wb.save(output)
    return output.getvalue()


# --- ФИНАЛ И ОТПРАВКА ---
st.divider()
if st.button("📊 Сформировать экспертный отчет", key="final_btn_report"):
    # Проверка обязательных полей
    mandatory_fields = [
        client_info['Город'], 
        client_info['Наименование компании'], 
        client_info['ФИО контактного лица'],
        client_info['Должность'],
        client_info['Контактный телефон']
    ]
    
    if not all(mandatory_fields):
        st.error("⚠️ Пожалуйста, заполните все обязательные поля шапки (отмечены *)!")
    elif not data:
        st.error("Технические данные не заполнены!")
    else:
        with st.spinner("Генерация отчета и отправка в Telegram..."):
            f_score = min(score, 100)
            report_bytes = make_excel(client_info, data, f_score)
            
            # --- ЛОГИКА ТЕЛЕГРАМ ---
            try:
                url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
                caption = (f"🛡️ *Новый аудит: {client_info['Наименование компании']}*\n\n"
                           f"📍 *Город:* {client_info['Город']}\n"
                           f"📊 *Зрелость:* {f_score}/100\n"
                           f"👤 *Контакт:* {client_info['ФИО контактного лица']}\n"
                           f"📞 *Тел:* {client_info['Контактный телефон']}\n"
                           f"📅 *Дата:* {client_info['Дата заполнения']}")
                
                files = {'document': (f"Audit_{client_info['Наименование компании']}_2026.xlsx", report_bytes)}
                payload = {"chat_id": CHAT_ID, "caption": caption, "parse_mode": "Markdown"}
                
                r = requests.post(url, data=payload, files=files)
                
                if r.ok:
                    st.success(f"Аналитика для {client_info['Наименование компании']} отправлена в группу!")
                    st.balloons()
                else:
                    st.error(f"Ошибка Telegram: {r.text}")
            except Exception as e:
                st.error(f"Сбой отправки: {e}")
            
            st.download_button(f"📥 Скачать Excel {client_info['Наименование компании']}", report_bytes, f"Audit_{client_info['Наименование компании']}.xlsx")

st.info("Разработано Ivan Rudoy | Khalil Audit | Almaty 2026")
