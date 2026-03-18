with col_h1:
    client_info['Город'] = st.text_input("Город:*")
    client_info['Наименование компании'] = st.text_input("Наименование компании:*")
    
    # 1. Поле сайта (откуда берем домен)
    site_input = st.text_input("Сайт компании (например, khalil.kz):", key="site_field")
    client_info['Сайт компании'] = site_input

    # Очищаем домен от протоколов и лишних знаков
    clean_domain = site_input.replace("https://", "").replace("http://", "").replace("www.", "").split('/')[0]

    # 2. Поле Email с жесткой фиксацией
    if clean_domain:
        # Используем колонки, чтобы текст "@домен" стоял сразу за полем ввода
        email_col1, email_col2 = st.columns([3, 2])
        with email_col1:
            email_prefix = st.text_input("Email контактного лица:", placeholder="info", key="email_prefix")
        with email_col2:
            # Выводим домен как обычный текст, который нельзя отредактировать
            st.markdown(f"<div style='padding-top: 25px; font-size: 18px; font-weight: bold; color: #1F4E78;'>@{clean_domain}</div>", unsafe_allow_html=True)
        
        # Фиксируем полный email
        client_info['Email'] = f"{email_prefix}@{clean_domain}" if email_prefix else ""
    else:
        st.warning("👈 Сначала введите сайт компании, чтобы зафиксировать домен почты")
        client_info['Email'] = ""
