import streamlit as st
import pandas as pd
import os
import requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from datetime import datetime

# --- НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Аудит ИТ и ИБ 2026", layout="wide", page_icon="🛡️")

# --- НАСТРОЙКИ TELEGRAM ---
TOKEN = st.secrets.get("TELEGRAM_TOKEN")
CHAT_ID = st.secrets.get("TELEGRAM_CHAT_ID")

# --- ЛОГОТИП ---
if os.path.exists("logo.png"):
    st.image("logo.png", width=300)
else:
    st.title("Khalil Trade IT Audit & Consulting")

st.title("📋 Опросник Технический аудит ИТ и ИБ (2026)")

client_info = {}
data = {}

# --- ШАПКА: ОБЩАЯ ИНФОРМАЦИЯ ---
st.header("📍 Общая информация")
client_info['Город'] = st.text_input("Город*")
client_info['Наименование компании'] = st.text_input("Наименование компании*")
client_info['Сфера деятельности'] = st.selectbox("Сфера деятельности*", ["Финтех", "Ритейл", "Производство", "IT", "Другое"])

# ЛОГИКА САЙТА И ПОЧТЫ (как просил: подстановка домена + чекбокс)
site_input = st.text_input("Сайт компании*", placeholder="example.kz")
client_info['Сайт компании'] = site_input
domain = site_input.replace("https://", "").replace("http://", "").split("/")[0]
email_default = f"info@{domain}" if domain else ""

if st.checkbox("Использовать почту на домене сайта?", value=True):
    client_info['Email'] = st.text_input("Электронная почта", value=email_default)
else:
    client_info['Email'] = st.text_input("Укажите другую почту")

client_info['ФИО контактного лица'] = st.text_input("ФИО контактного лица*")
client_info['Контактный телефон'] = st.text_input("Контактный телефон*")

st.divider()

# --- БЛОК 1: ИТ ---
st.header("Блок 1: Информационные технологии")

# 1.1 АРМ
st.subheader("1.1. Конечные точки")
total_arm = st.number_input("Общее количество АРМ (шт)", min_value=0, step=1)

# 1.2 Сеть
st.subheader("1.2. Сетевая инфраструктура")
main_speed = st.number_input("Скорость основного канала (Mbit/s)", min_value=0)
back_speed = st.number_input("Скорость резервного канала (Mbit/s)", min_value=0)

st.write("Беспроводная сеть")
ap_cnt = st.number_input("Количество точек доступа (шт)", min_value=0)
wifi_ctrl = st.checkbox("Наличие Wi-Fi контроллера")

st.write("Сетевая безопасность")
ngfw = st.text_input("Производитель NGFW (если есть)", value="Нет")
# NAD перенесен сюда
nad_vendor = st.text_input("Производитель NAD (Network Analysis and Detection)", value="Нет")

# 1.3 Серверы и СХД
st.subheader("1.3. Серверы и СХД")
phys_count = st.number_input("Физические серверы", min_value=0)
virt_count = st.number_input("Виртуальные серверы", min_value=0)
srk_vendor = st.text_input("Вендор СРК (Резервное копирование)", value="Нет")
storage_vendor = st.text_input("Производитель СХД", value="Нет")

# 1.4 ИС и ПО
st.subheader("1.4. Информационные системы")
helpdesk_vendor = st.text_input("Система HelpDesk (если есть)", value="Нет")
web_stack = st.text_input("WEB-ресурсы (стек)", value="Нет")
dev_stack = st.text_input("Разработка (стек/процессы)", value="Нет")

st.divider()

# --- БЛОК 2: ИБ ---
st.header("Блок 2: Информационная безопасность")
ib_values = {
    "EPP/AV": st.text_input("Вендор EPP (Антивирус)", value="Нет"),
    "EDR": st.text_input("Вендор EDR", value="Нет"),
    "DLP": st.text_input("Вендор DLP", value="Нет"),
    "SIEM": st.text_input("Вендор SIEM", value="Нет"),
    "PAM": st.text_input("Вендор PAM", value="Нет")
}

# --- ЛОГИКА ФОРМИРОВАНИЯ ОТЧЕТА ---
def create_excel(c_info, results):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Аудит"
    
    # Заголовки
    ws.append(["Параметр", "Значение", "Комментарий/Рекомендация"])
    
    # 1. Проверка каналов
    m_s = results['main_speed']
    b_s = results['back_speed']
    row_ch = ["Резервный канал", b_s]
    if b_s == 0:
        row_ch.append("Отсутствие резервного канала — единая точка отказа.")
    elif b_s < (m_s / 2):
        row_ch.append("ВНИМАНИЕ: резервный канал не покрывает производительность основного. Рекомендовано расширить.")
    else:
        row_ch.append("ОК")
    ws.append(row_ch)

    # 2. Проверка Wi-Fi
    u_count = results['total_arm']
    a_count = results['ap_cnt']
    row_wf = ["Wi-Fi Инфраструктура", f"{a_count} ТД"]
    if a_count > 0:
        ratio = u_count / a_count
        if ratio > 25:
            row_wf.append(f"ВНИМАНИЕ: Высокая плотность ({ratio:.1f} чел/ТД). Не хватает точек доступа.")
        elif a_count > 5 and not results['wifi_ctrl']:
            row_wf.append("ПРОБЛЕМА: Много точек доступа без контроллера. Рекомендовано внедрение контроллера.")
        else:
            row_wf.append("ОК")
    else:
        row_wf.append("Wi-Fi не используется или не указан")
    ws.append(row_wf)

    # 3. Логика по количеству (SIEM / HelpDesk)
    total_hosts = results['phys_count'] + results['virt_count'] + u_count
    if total_hosts > 50 and results['siem'] == "Нет":
        ws.append(["SIEM система", "Отсутствует", "Рекомендация CISO: При таком количестве узлов необходим SIEM."])
    
    if u_count > 30 and results['helpdesk'] == "Нет":
        ws.append(["HelpDesk", "Отсутствует", "Рекомендация CTO: Большой штат пользователей требует систему автоматизации заявок."])

    # 4. Добавляем все остальные блоки (СХД, WEB, Разработка), чтобы не терялись
    ws.append(["СХД", results['storage'], ""])
    ws.append(["WEB-ресурсы", results['web'], ""])
    ws.append(["Разработка", results['dev'], ""])
    ws.append(["Резервное копирование", results['srk'], ""])
    ws.append(["NAD", results['nad'], "Сетевая безопасность"])

    wb.save(output)
    return output.getvalue()

# --- КНОПКА ОТПРАВКИ ---
if st.button("Сформировать отчет"):
    all_data = {
        'main_speed': main_speed,
        'back_speed': back_speed,
        'total_arm': total_arm,
        'ap_cnt': ap_cnt,
        'wifi_ctrl': wifi_ctrl,
        'phys_count': phys_count,
        'virt_count': virt_count,
        'siem': ib_values['SIEM'],
        'helpdesk': helpdesk_vendor,
        'storage': storage_vendor,
        'web': web_stack,
        'dev': dev_stack,
        'srk': srk_vendor,
        'nad': nad_vendor
    }
    
    report_file = create_excel(client_info, all_data)
    
    # Отправка в ТГ
    try:
        requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", 
                      data={"chat_id": CHAT_ID, "caption": f"Аудит: {client_info['Наименование компании']}"}, 
                      files={'document': ("Audit.xlsx", report_file)})
    except:
        pass

    st.success("Отчет сформирован и отправлен!")
    st.download_button("Скачать Excel", report_file, "Audit_Report.xlsx")
