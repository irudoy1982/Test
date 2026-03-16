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
client_info = {}
score = 0

# --- ШАПКА: ИНФОРМАЦИЯ О КОМПАНИИ ---
st.header("🏢 Общая информация о компании")
col_c1, col_c2 = st.columns(2)

with col_c1:
    client_info['Наименование компании'] = st.text_input("Наименование компании:")
    client_info['Сайт Компании'] = st.text_input("Сайт Компании:")
    client_info['Контактное лицо'] = st.text_input("Контактное лицо (ФИО):")

with col_c2:
    client_info['Должность'] = st.text_input("Должность:")
    client_info['Контактный email'] = st.text_input("Контактный email:")
    client_info['Контактный телефон'] = st.text_input("Контактный телефон:")

st.divider()

# --- БЛОК 1: ТЕХНИЧЕСКАЯ ИНФОРМАЦИЯ ---
st.header("Блок 1: Инфраструктура")

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
    data['1.2. Физические серверы'] = st.number_input("Количество физических серверов:", min_value=0, step=1)
with col_s2:
    data['1.2. Виртуальные серверы'] = st.number_input("Количество виртуальных серверов:", min_value=0, step=1)

# Блок ИБ (упрощенно для примера, сохраняя вашу логику)
st.header("Блок 2: Информационная Безопасность")
ib_list = {"Резервное копирование": 20, "DLP (Защита от утечек)": 15, "EDR/Antimalware": 15}
for label, pts in ib_list.items():
    if st.checkbox(label):
        v = st.text_input(f"Вендор {label}:", key=f"v_{label}")
        data[label] = f"Да ({v if v else 'не указан'})"
        score += pts
    else:
        data[label] = "Нет"

# --- ЛОГИКА РЕКОМЕНДАЦИЙ ---
def get_recommendation(key, value):
    db = {
        "Нет": "Критический риск. Отсутствие данной системы снижает прозрачность и безопасность ИТ-инфраструктуры.",
        "Резервное копирование": "КРИТИЧНО: Рекомендуется стратегия 3-2-1. Без бэкапов данные под угрозой.",
        "EDR/Antimalware": "Рекомендуются современные системы защиты конечных точек (EDR) для борьбы со сложными угрозами."
    }
    if "Нет" in str(value) or value == 0:
        return db.get(key, db["Нет"])
    return "Конфигурация соответствует базовым стандартам."

# --- ЭКСЕЛЬ ГЕНЕРАЦИЯ ---
def make_excel(c_info, results, final_score):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Анализ Аудита"

    # Стили
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # 1. Заголовок
    ws.merge_cells('A1:D1')
    ws['A1'] = "ЭКСПЕРТНЫЙ ОТЧЕТ ПО ТЕХНИЧЕСКОМУ АУДИТУ"
    ws['A1'].alignment = Alignment(horizontal='center')
    ws['A1'].font = Font(bold=True, size=14)

    # 2. Информация о клиенте
    ws['A3'] = "ИНФОРМАЦИЯ О КОМПАНИИ"
    ws['A3'].font = Font(bold=True)
    
    current_row = 4
    for k, v in c_info.items():
        ws.cell(row=current_row, column=1, value=k).font = Font(bold=True)
        ws.cell(row=current_row, column=2, value=str(v))
        current_row += 1

    ws.cell(row=current_row, column=1, value="ИНДЕКС ЗРЕЛОСТИ:").font = Font(bold=True)
    ws.cell(row=current_row, column=2, value=f"{final_score} / 100")
    current_row += 2

    # 3. Таблица результатов
    headers = ["Параметр", "Значение", "Статус", "Рекомендация эксперта"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=current_row, column=i, value=h)
        cell.fill = header_fill
        cell.font = white_font

    current_row += 1
    for k, v in results.items():
        ws.cell(row=current_row, column=1, value=k).border = border
        ws.cell(row=current_row, column=2, value=str(v)).border = border
        
        rec = get_recommendation(k, v)
        status = "РИСК" if "Нет" in str(v) or "Критический" in rec else "ОК"
        
        st_cell = ws.cell(row=current_row, column=3, value=status)
        st_cell.border = border
        if status == "РИСК":
            st_cell.font = Font(color="FF0000", bold=True)
            
        rc_cell = ws.cell(row=current_row, column=4, value=rec)
        rc_cell.border = border
        rc_cell.alignment = Alignment(wrap_text=True)
        current_row += 1

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 50

    # Логотип
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
if st.button("📊 Сформировать экспертный отчет"):
    if not client_info['Наименование компании']:
        st.error("Пожалуйста, заполните наименование компании!")
    else:
        f_score = min(score, 100)
        report = make_excel(client_info, data, f_score)
        st.success(f"Отчет для {client_info['Наименование компании']} готов!")
        st.download_button("📥 Скачать экспертный Excel", report, f"Audit_{client_info['Наименование компании']}_2026.xlsx")

st.info("Разработано Ivan Rudoy. По вопросам системной интеграции — звоните!")
