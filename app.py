
                current.append(searched_group)
                st.session_state.selected_groups = current

    selectable_groups = [
        "A1", "A2", "A3",
        "B1", "B2", "B3",
        "C1", "C2", "C3",
        "D1", "D2", "D3",
        "E1", "E2", "E3",
        "F1", "F2", "F3",
        "G1", "G2", "G3",
        "H1", "H2", "H3",
        "K1", "K2", "K3",
        "DİĞER",
    ]

    # Yalnızca gerçekten mevcut grupları seçilebilir listeye al.
    available_groups = [
        group_name
        for group_name in selectable_groups
        if pharmacies["Grup"].eq(group_name).any()
    ]

    if "selected_groups" not in st.session_state:
        st.session_state.selected_groups = available_groups.copy()

    # Dosya/grup listesi değiştiğinde session_state içindeki eski değerleri temizle.
    st.session_state.selected_groups = [
        group_name
        for group_name in st.session_state.selected_groups
        if group_name in available_groups
    ]

    st.sidebar.markdown("### Alt grup seçimi")

    col_select_all, col_clear_all = st.sidebar.columns(2)

    with col_select_all:
        if st.button(
            "Tümünü seç",
            use_container_width=True,
            key="select_all_groups",
        ):
            st.session_state.selected_groups = available_groups.copy()
            st.rerun()

    with col_clear_all:
        if st.button(
            "Temizle",
            use_container_width=True,
            key="clear_all_groups",
        ):
            st.session_state.selected_groups = []
            st.rerun()

    selected_groups = st.sidebar.multiselect(
        "Haritada gösterilecek alt gruplar",
        options=available_groups,
        key="selected_groups",
        placeholder="Alt grup seçin",
    )

    st.sidebar.caption(
        f"Seçili grup: {len(selected_groups)} / {len(available_groups)}"
    )

    # Grup sayılarını kompakt şekilde göster.
    with st.sidebar.expander("Grup adetleri", expanded=False):
        for group_name in available_groups:
            group_label = "Diğer" if group_name == "DİĞER" else group_name
            st.metric(
                group_label,
                int(pharmacies["Grup"].eq(group_name).sum()),
            )

    # Haritada sadece kullanıcının seçtiği gruplar gösterilir.
    visible_pharmacies = pharmacies.loc[
        pharmacies["Grup"].isin(selected_groups)
    ].copy()

    st.sidebar.metric(
        "Haritada görünen eczane",
        len(visible_pharmacies),
    )

    st.sidebar.caption(
        "Seçilmeyen alt grupların hem noktaları hem bağlantı çizgileri haritadan kaldırılır."
    )

    # --------------------------------------------------------
    # ANA EKRAN
    # --------------------------------------------------------

    st.title(
        "Denizli Eczane Haritası — Grup A + B + C + D + E + F + G + H + K"
    )

    st.caption(
        "A1-A3, B1-B3, C1-C3, D1-D3, E1-E3, F1-F3, G1-G3, H1-H3 ve K1-K3 kendi içlerinde kesikli çizgilerle bağlanır · "
        "diğer eczaneler şimdilik düz/nötr nokta olarak kalır · "
        "kırmızı yoğunluk çemberi tüm eczaneleri sayar"
    )

    if visible_pharmacies.empty:
        st.warning(
            "Haritada göstermek için en az bir alt grup seçin."
        )
        st.stop()

    m = build_map(
        visible_pharmacies,
        searched_pharmacy=searched_pharmacy,
    )

    components.html(
        m.get_root().render(),
        height=900,
        scrolling=False,
    )


except Exception as exc:

    st.exception(exc)
