# =========================================================
# TAMPILAN UNTUK SUPPLIER (AKSES VIA LINK TOKEN & KATEGORI)
# =========================================================
query_params = st.query_params
token = query_params.get("token", None)
kat_param = query_params.get("kat", None)

if token:
    # 1. CARI DATA SUPPLIER BERDASARKAN TOKEN
    df_sup = st.session_state.get("df_supplier_state", pd.DataFrame())
    
    if not df_sup.empty and "TOKEN" in df_sup.columns:
        supplier_row = df_sup[df_sup["TOKEN"] == token]
    else:
        supplier_row = pd.DataFrame()

    if supplier_row.empty:
        st.error("❌ Token tidak valid atau supplier tidak ditemukan.")
    else:
        nama_supplier = supplier_row.iloc[0].get("NAMA SUPPLIER", "Supplier")
        kode_supplier = supplier_row.iloc[0].get("KODE SUPPLIER", "-")

        # 2. FILTER KATEGORI BARANG
        if kat_param and kat_param.upper() != "ALL":
            kategori_list = [k.replace("_", " ").strip().upper() for k in kat_param.split(",")]
        else:
            kategori_list = []

        # Header Form Penawaran
        st.markdown("## **FORM PENAWARAN HARGA**")
        
        # Info Box Supplier & Kategori
        kat_display = ", ".join(kategori_list) if kategori_list else "Semua Kategori"
        st.info(f"**Supplier:** {kode_supplier} - {nama_supplier}\n\n**Jenis BB:** {kat_display}")

        # 3. AMBIL DATA ITEM BAHAN BAKU BERDASARKAN KATEGORI
        df_master_bb = st.session_state.get("df_bb_state", pd.DataFrame())

        if not df_master_bb.empty and "JENIS BB" in df_master_bb.columns:
            if kategori_list:
                df_filtered = df_master_bb[df_master_bb["JENIS BB"].str.upper().isin(kategori_list)].copy()
            else:
                df_filtered = df_master_bb.copy()
        else:
            df_filtered = pd.DataFrame()

        if df_filtered.empty:
            st.warning("⚠️ Tidak ada daftar item bahan baku yang sesuai dengan kategori ini.")
        else:
            st.write("Silakan isi kolom **HARGA PENAWARAN** di bawah ini dengan angka saja (tanpa titik/rp):")

            # Reset Index & Urutan Kolom
            df_filtered = df_filtered.reset_index(drop=True)
            df_filtered["NO"] = df_filtered.index + 1
            
            # Buat kolom Form Input Harga
            with st.form("form_penawaran_supplier"):
                # Header Tabel Kustom
                col_no, col_pn, col_jenis, col_item, col_sat, col_harga = st.columns([0.6, 1.8, 1.2, 3, 1, 2])
                col_no.markdown("**NO**")
                col_pn.markdown("**P/N**")
                col_jenis.markdown("**JENIS BB**")
                col_item.markdown("**ITEM BB**")
                col_sat.markdown("**SAT**")
                col_harga.markdown("**HARGA PENAWARAN**")
                st.markdown("---")

                input_harga_dict = {}

                # Loop Setiap Item untuk Membuat Baris Input
                for idx, row in df_filtered.iterrows():
                    c_no, c_pn, c_jenis, c_item, c_sat, c_harga = st.columns([0.6, 1.8, 1.2, 3, 1, 2])
                    
                    pn_val = str(row.get("P/N", "-"))
                    jenis_val = str(row.get("JENIS BB", "-"))
                    item_val = str(row.get("ITEM BB", "-"))
                    sat_val = str(row.get("SATUAN", row.get("SAT", "-")))

                    c_no.write(f"{idx + 1}")
                    c_pn.write(pn_val)
                    c_jenis.write(jenis_val)
                    c_item.write(item_val)
                    c_sat.write(sat_val)
                    
                    # Field Input Harga per Baris
                    input_harga_dict[pn_val] = c_harga.number_input(
                        label=f"Harga {pn_val}",
                        min_value=0,
                        step=500,
                        value=0,
                        key=f"harga_{pn_val}_{idx}",
                        label_visibility="collapsed"
                    )

                st.markdown("<br>", unsafe_allow_html=True)
                submit_btn = st.form_submit_button("💾 Simpan Penawaran Harga", type="primary", use_container_width=True)

                if submit_btn:
                    # PROSES SIMPAN KE DATABASE / SESSION STATE
                    data_simpan = []
                    for pn_key, harga_val in input_harga_dict.items():
                        if harga_val > 0:
                            data_simpan.append({
                                "KODE SUPPLIER": kode_supplier,
                                "NAMA SUPPLIER": nama_supplier,
                                "P/N": pn_key,
                                "HARGA": harga_val,
                                "TANGGAL": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")
                            })
                    
                    if data_simpan:
                        # Tambahkan ke database hasil penawaran
                        df_hasil = st.session_state.get("df_penawaran_state", pd.DataFrame())
                        st.session_state["df_penawaran_state"] = pd.concat([df_hasil, pd.DataFrame(data_simpan)], ignore_index=True)
                        st.success("✅ Berhasil menyimpan data penawaran harga! Terima kasih.")
                    else:
                        st.warning("⚠️ Mohon isi minimal 1 harga item bahan baku sebelum menyimpan.")
