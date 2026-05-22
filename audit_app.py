if st.button("Сформировать экспертный отчет"):
        # === 1. СОЗДАНИЕ КОНТЕЙНЕРОВ ДЛЯ АНИМАЦИИ И НАДПИСЕЙ ===
        status_header = st.empty()
        warning_placeholder = st.empty()
        log_matrix_placeholder = st.empty()  # Сюда будут динамически выводиться 3 строки логов
        
        # Выводим предупреждение и статус
        warning_placeholder.warning("⚠️ ВНИМАНИЕ: Выполняется сложный анализ матрицы угроз. Пожалуйста, не закрывайте вкладку...")
        status_header.markdown("### 📊 Статус: *Кросс-табличный анализ рисков...*")

        # Шаги для скроллинга логов
        steps = [
            "Расчет базовых технологических индексов и весов уязвимостей...",
            "Валидация введенных данных на соответствие комплаенс-метрикам ISO 27001 / NIST CSF...",
            "Сопоставление ИТ-ландшафта с отраслевой матрицей угроз и расчет рекомендаций...",
            "Выполнение математических вычислений и автоматический подбор тех. стека...",
            "Формирование структуры книги Excel и генерация динамических таблиц...",
            "Применение корпоративного стиля Khalil Consulting: калибровка ячеек..."
        ]

        # === 2. ЦИКЛ ДИНАМИЧЕСКОГО СКРОЛЛИНГА ЛОГОВ (СТРОГО 3 СТРОКИ) ===
        for i in range(len(steps)):
            current_time = time.strftime('%H:%M:%S')
            log_content = "<div style='font-family: monospace; line-height: 1.6;'>"
            
            # Строка 1: Прошлый шаг (зеленый)
            if i > 0:
                log_content += f"<div style='color: #2e7d32; opacity: 0.6;'>✅ `[{current_time}]` {steps[i-1]}</div>"
            else:
                log_content += "<div style='color: #666; opacity: 0.3;'>... ожидание запуска системы ...</div>"
                
            # Строка 2: Текущий активный шаг (оранжевый флэш)
            log_content += f"<div style='color: #f57c00; font-weight: bold; margin: 4px 0;'>🔄 `[{current_time}]` {steps[i]}</div>"
            
            # Строка 3: Следующий шаг в очереди (серый)
            if i < len(steps) - 1:
                log_content += f"<div style='color: #757575; opacity: 0.5;'>⏳ `[очередь]` {steps[i+1]}</div>"
            else:
                log_content += "<div style='color: #00ff66; font-weight: bold;'>🚀 Финализация структуры файла...</div>"
                
            log_content += "</div>"
            
            # Обновляем контейнер (создается эффект скроллинга)
            log_matrix_placeholder.markdown(log_content, unsafe_allow_html=True)
            time.sleep(0.8)

        # === 3. РАСЧЕТ СКОРИНГА И РЕАЛЬНАЯ ГЕНЕРАЦИЯ EXCEL ===
        f_score = min(score, 100)
        report_bytes = make_expert_excel(client_info, data, f_score)

        # === 4. ОТПРАВКА В TELEGRAM ===
        try:
            comp_name = client_info.get('Наименование компании') or 'Не указана'
            
            contact_parts = []
            for key in ['ФИО контактного лица', 'Должность', 'Контактный телефон', 'Email']:
                val = client_info.get(key)
                if val:
                    contact_parts.append(str(val))
            contact_str = " | ".join(contact_parts) if contact_parts else "Не указаны"

            tg_message = (
                f"🔔 *Сгенерирован экспертный отчет!*\n\n"
                f"🏢 *Компания:* {comp_name}\n"
                f"👤 *Контакты:* {contact_str}\n"
                f"🛡️ *Итоговый скоринг:* {f_score}%\n"
                f"📅 *Дата:* {time.strftime('%d.%m.%Y %H:%M:%S')}"
            )
            
            if TOKEN and CHAT_ID:
                requests.post(f"https://api.telegram.org/bot{TOKEN}/sendMessage", data={"chat_id": CHAT_ID, "text": tg_message, "parse_mode": "Markdown"}, timeout=5)
                requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", data={"chat_id": CHAT_ID, "caption": f"📁 Отчет: {comp_name}"}, files={'document': (f"Audit_v10_{comp_name.replace(' ', '_')}.xlsx", report_bytes)}, timeout=10)
        except Exception as tg_err:
            print(f"Ошибка Telegram: {tg_err}")

        # === 5. МГНОВЕННАЯ ОЧИСТКА ЭКРАНА И ФИНАЛЬНЫЙ БАННЕР ===
        # Схлопываем логи и варнинг, чтобы освободить место под баннер выгрузки
        log_matrix_placeholder.empty()
        warning_placeholder.empty()
        
        status_header.markdown("### ✅ Статус: *Экспертный анализ успешно завершен!*")
        
        import base64
        company_filename = comp_name.replace(" ", "_")
        b64_report = base64.b64encode(report_bytes).decode('utf-8')

        st.markdown(f"""
        <style>
            .cyber-monolith-banner {{
                background-color: #0e1117;
                border: 2px solid #00ff66;
                border-radius: 8px;
                padding: 30px 25px;
                text-align: center;
                box-shadow: 0px 0px 20px rgba(0, 255, 102, 0.15);
                font-family: 'Courier New', monospace;
                margin-top: 20px;
                margin-bottom: 25px;
                width: 100%;
                box-sizing: border-box;
            }}
            .cyber-title {{ color: #00ff66; margin: 0; font-size: 24px; letter-spacing: 2px; font-weight: bold; }}
            .cyber-subtitle {{ color: #666; font-size: 11px; margin-top: 6px; margin-bottom: 25px; letter-spacing: 1px; }}
            .cyber-download-link {{
                display: block; width: 100%; box-sizing: border-box; background-color: rgba(0, 255, 102, 0.04);
                color: #ffffff !important; border: 1px dashed #00ff66; border-radius: 4px; padding: 15px 20px;
                font-weight: bold; font-size: 13px; text-decoration: none; transition: all 0.25s ease;
            }}
            .cyber-download-link:hover {{ background-color: rgba(0, 255, 102, 0.16) !important; color: #00ff66 !important; border: 1px solid #00ff66; }}
        </style>
        
        <div class="cyber-monolith-banner">
            <h1 class="cyber-title">🛡️ SECURITY AUDIT COMPLETE</h1>
            <p class="cyber-subtitle">STATUS CODE: 200 SUCCESS | CORE V10.5</p>
            <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_report}" 
               download="Audit_v10_{company_filename}.xlsx" 
               class="cyber-download-link">
                🔒 ЭКСПЕРТНЫЙ ОТЧЕТ СКОМПИЛИРОВАН И ГОТОВ К ВЫГРУЗКЕ (XLSX)
            </a>
        </div>
        """, unsafe_allow_html=True)

        st.success(f"✔️ Экспертный анализ успешно завершен. Итоговый уровень защищенности: {f_score}%")
    
# Подвал приложения (БЕЗ отступов у левого края файла)
st.info("Khalil Audit System v10.5 | Ivan Rudoy Production | Almaty 2026")
