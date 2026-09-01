# ---------------------------------------------------------
# MODUL 4: WA & PO GENERATOR (FIXED EDIT PESAN & LINK ERROR)
# ---------------------------------------------------------
elif menu == "💬 WA & PO Generator":
    import urllib.parse

    st.subheader("💬 WA Message & PO Link Generator")
    st.caption("Kelola kategori barang yang disuplai oleh masing-masing vendor dan buat draft pesan WhatsApp beserta link penawaran harganya.")

    df_sup = st.session_state["df_supplier_state"]

    if df_sup.empty:
        st.warning("⚠️ Data supplier masih kosong. Silakan isi atau import data di menu '🤝 Data Supplier & Link Form' terlebih dahulu.")
    else:
        # Master list kategori
        LIST_KATEGORI = [
            "Ayam", "Beras", "Buah", "Cookies", "Daging", "Ikan", 
            "Keju", "Olahan", "Sayur", "Sembako", "Susu", "Tahu", 
            "Tempe", "Telur Ayam", "Telur Bebek", "Telur Puyuh"
        ]

        # 1. PILIH SUPPLIER
        st.markdown("##### 1️⃣ **Pilih Supplier & Atur Kategori Barang**")
        
        list_supplier_display = [f"{r['KODE SUPPLIER']} - {r['NAMA SUPPLIER']}" for _, r in df_sup.iterrows()]
        selected_sup_str = st.selectbox("Pilih Supplier:", list_supplier_display)
        
        idx_sup = list_supplier_display.index(selected_sup_str)
        row_sup = df_sup.iloc[idx_sup]

        # 2. EDIT KATEGORI SUPPLIER
        st.markdown("---")
        st.write(f"Edit kategori bahan baku yang disuplai oleh **{row_sup['NAMA SUPPLIER']}**:")

        kat_existing = []
        for col_bb in ["JENIS BB 1", "JENIS BB 2", "JENIS BB 3", "JENIS BB 4", "JENIS BB 5", "JENIS BB 6"]:
            val = row_sup.get(col_bb, "-")
            if pd.notna(val) and str(val).strip() not in ["-", "", "nan", "None"]:
                matched = next((k for k in LIST_KATEGORI if k.upper() == str(val).strip().upper()), None)
                if matched:
                    kat_existing.append(matched)

        selected_kategori = st.multiselect(
            "Pilih Kategori Barang (Maksimal 6 Jenis BB):",
            options=LIST_KATEGORI,
            default=kat_existing,
            max_selections=6,
            key=f"ms_kategori_{idx_sup}"
        )

        col_save, _ = st.columns([1, 3])
        with col_save:
            if st.button("💾 Simpan Perubahan Kategori", type="primary", use_container_width=True):
                for i in range(1, 7):
                    col_name = f"JENIS BB {i}"
                    val_to_save = selected_kategori[i-1] if i <= len(selected_kategori) else "-"
                    st.session_state["df_supplier_state"].at[idx_sup, col_name] = val_to_save
                
                st.success(f"Kategori untuk {row_sup['NAMA SUPPLIER']} berhasil diperbarui!")
                st.rerun()

        # 3. GENERATE LINK & PESAN WA DENGAN FITUR EDIT MANUAL
        st.markdown("---")
        st.markdown("##### 2️⃣ **Preview & Edit Pesan WhatsApp**")

        token = row_sup.get("TOKEN", "")
        kat_param = ",".join([k.replace(" ", "_") for k in selected_kategori]) if selected_kategori else "ALL"
        
        # Buat link form supplier
        link_form_khusus = f"{WEB_APP_URL}?token={token}&kat={kat_param}"

        no_wa = str(row_sup.get("NO WA", "")).replace("+", "").replace(" ", "").replace("-", "").split(".")[0]
        nama_supplier = row_sup.get("NAMA SUPPLIER", "Bapak/Ibu Vendor")
        kategori_txt = ", ".join(selected_kategori) if selected_kategori else "Bahan Baku"

        # Template default
        default_pesan = (
            f"Halo *{nama_supplier}*,\n\n"
            f"Kami dari tim Procurement Koperasi YK/SPPG. "
            f"Mohon untuk mengisi update harga penawaran harian/mingguan untuk kategori *[{kategori_txt}]* "
            f"melalui link form resmi berikut ini:\n\n"
            f"🔗 {link_form_khusus}\n\n"
            f"Terima kasih atas kerja samanya."
        )

        # Text area dengan KEY agar editan user tersimpan di session_state
        pesan_edited = st.text_area(
            "Silakan edit draft pesan di bawah ini jika diperlukan:", 
            value=default_pesan, 
            height=180,
            key=f"txt_wa_{idx_sup}"
        )

        col_wa1, col_wa2 = st.columns([1, 1])
        with col_wa1:
            # Menggunakan pesan_edited (hasil editan user) bukan default_pesan
            encoded_msg = urllib.parse.quote(pesan_edited)
            wa_url = f"https://wa.me/{no_wa}?text={encoded_msg}"
            
            st.markdown(
                f'<a href="{wa_url}" target="_blank" style="text-decoration:none;">'
                f'<div style="background-color:#25D366; color:white; padding:12px; border-radius:8px; text-align:center; font-weight:bold;">'
                f'📲 Kirim Hasil Editan via WhatsApp (+{no_wa})</div></a>',
                unsafe_allow_html=True
            )
            
        with col_wa2:
            st.info(f"💡 Apabila kamu mengubah isi teks di atas, tombol WhatsApp akan otomatis mengirimkan teks hasil editan terbaru kamu.")
