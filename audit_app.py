# 1.4 СХД
st.write("---")
st.subheader("1.4. Системы хранения данных (СХД)")

if st.toggle("Есть собственная СХД", key="storage_toggle"):
    st_media_sel = st.multiselect(
        "Типы носителей",
        ["HDD (NL-SAS / SATA)", "SSD (SATA / SAS)", "NVMe", "SCM"],
        key="st_media"
    )
    data['1.4.1. Типы носителей'] = ", ".join(st_media_sel)

    col_pct1, col_pct2 = st.columns(2)

    with col_pct1:
        cnt_hdd = st.number_input(
            "Количество дисков HDD",
            min_value=0,
            step=1,
            key="cnt_hdd"
        )
        data['1.4.2. Кол-во HDD'] = cnt_hdd

    with col_pct2:
        cnt_ssd = st.number_input(
            "Количество дисков SSD",
            min_value=0,
            step=1,
            key="cnt_ssd"
        )
        data['1.4.3. Кол-во SSD'] = cnt_ssd

    if st_media_sel and (cnt_hdd + cnt_ssd == 0):
        st.info("ℹ️ Укажите количество дисков для СХД.")

    col_chk1, col_chk2 = st.columns(2)

    with col_chk1:
        data['1.4.4. Гибридная СХД'] = st.checkbox(
            "Используется гибридная СХД",
            key="hybrid_st"
        )

    with col_chk2:
        data['1.4.5. All-Flash'] = st.checkbox(
            "Есть All-Flash массивы",
            key="allflash_st"
        )

    data['1.4.6. RAID-группы'] = ", ".join(
        st.multiselect(
            "Используемые RAID-группы",
            ["RAID 0", "RAID 1", "RAID 5", "RAID 6", "RAID 10", "RAID 50", "RAID 60", "JBOD"],
            key="raid_list"
        )
    )

    data['1.4. Примечание'] = st.text_area(
        "Примечание к разделу 1.4",
        placeholder="Напр.: используется SAN + tiering, планируется переход на NVMe",
        key="note_1_4"
    )
