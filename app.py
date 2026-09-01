# --- PILIHAN MENU DARI SIDEBAR ---
if menu == "📊 Dashboard & HET":
    st.subheader("📊 Dashboard Utama & Pemantauan HET")
    # ... (kode modul dashboard) ...

elif menu == "🏬 Kelola Data Dapur":
    st.subheader("🏬 Manajemen Data Dapur SPPG & Peta Sebaran")
    # ... (kode modul dapur) ...

elif menu == "🤝 Data Supplier & Link Form":
    st.subheader("🤝 Kelola Data Supplier & Link Web App")
    
    # --- FITUR UPLOAD FILE EXCEL / SPREADSHEET ---
    with st.expander("📂 **Upload Data Supplier dari File Excel / Google Sheets**", expanded=True):
        st.write("Unggah file `.xlsx`, `.xls`, atau `.csv` yang berisi daftar supplier kamu.")
        uploaded_file = st.file_uploader("Pilih file Excel Supplier:", type=["xlsx", "xls", "xlsb", "csv"])
        
        if uploaded_file is not None:
            try:
                # Membaca file
                if uploaded_file.name.endswith('.csv'):
                    df_upload = pd.read_csv(uploaded_file)
                else:
                    df_upload = pd.read_excel(uploaded_file)
                
                # Deteksi baris header jika terdapat judul/empty row di atasnya
                if "Supp Code" not in df_upload.columns and "Supplier Name" not in df_upload.columns:
                    for i in range(min(10, len(df_upload))):
                        row_vals = df_upload.iloc[i].astype(str).tolist()
                        if any("SUPP CODE" in str(v).upper() or "SUPPLIER NAME" in str(v).upper() for v in row_vals):
                            df_upload.columns = df_upload.iloc[i]
                            df_upload = df_upload.iloc[i+1:].reset_index(drop=True)
                            break

                # Cleaning nama kolom
                df_upload.columns = [str(c).strip() for c in df_upload.columns]
                
                # Mapping Nama Kolom
                mapped_df = pd.DataFrame()
                
                # 1. Kode Supplier
                col_code = next((c for c in df_upload.columns if 'SUPP CODE' in c.upper() or 'KODE' in c.upper()), None)
                mapped_df["KODE SUPPLIER"] = df_upload[col_code].astype(str) if col_code else [f"SUP-{i+1:03d}" for i in range(len(df_upload))]
                
                # 2. Nama Supplier
                col_name = next((c for c in df_upload.columns if 'SUPPLIER NAME' in c.upper() or 'NAMA' in c.upper()), None)
                mapped_df["NAMA SUPPLIER"] = df_upload[col_name].astype(str) if col_name else "Tanpa Nama"
                
                # 3. Nomor WA / Phone (Auto-format ke format 628xxx)
                col_phone = next((c for c in df_upload.columns if 'PHONE' in c.upper() or 'WA' in c.upper() or 'TELP' in c.upper()), None)
                if col_phone:
                    def clean_phone(p):
                        p_str = str(p).replace("-", "").replace(" ", "").replace("+", "").split(".")[0].strip()
                        if p_str.startswith("0"):
                            return "62" + p_str[1:]
                        elif not p_str.startswith("62") and len(p_str) > 5:
                            return "62" + p_str
                        return p_str
                    mapped_df["NO WA"] = df_upload[col_phone].apply(clean_phone)
                else:
                    mapped_df["NO WA"] = "6285169796974"

                # 4. Kategori / Supply BB
                col_bb = next((c for c in df_upload.columns if 'SUPPLY BB' in c.upper() or 'JENIS' in c.upper() or 'KATEGORI' in c.upper()), None)
                mapped_df["JENIS BB 1"] = df_upload[col_bb].astype(str) if col_bb else "-"
                mapped_df["JENIS BB 2"] = "-"
                mapped_df["JENIS BB 3"] = "-"

                # 5. Token Unik & Link Form
                mapped_df["TOKEN"] = [buat_token() for _ in range(len(mapped_df))]
                mapped_df["LINK FORM"] = mapped_df["TOKEN"].apply(lambda t: f"{WEB_APP_URL}?token={t}")

                # Kolom Tambahan (PIC & Alamat)
                col_pic = next((c for c in df_upload.columns if 'PIC' in c.upper()), None)
                mapped_df["PIC"] = df_upload[col_pic].astype(str) if col_pic else "-"
                
                col_alamat = next((c for c in df_upload.columns if 'ALAMAT' in c.upper()), None)
                mapped_df["ALAMAT"] = df_upload[col_alamat].astype(str) if col_alamat else "-"

                mapped_df = mapped_df.dropna(subset=["NAMA SUPPLIER"]).reset_index(drop=True)

                st.success(f"✅ Berhasil membaca **{len(mapped_df)} data supplier** dari file!")
                st.dataframe(mapped_df[["KODE SUPPLIER", "NAMA SUPPLIER", "NO WA", "JENIS BB 1", "TOKEN"]].head(), use_container_width=True)

                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    if st.button("📥 Import & Timpa Data Supplier Baru", type="primary", use_container_width=True):
                        st.session_state["df_supplier_state"] = mapped_df
                        st.success("Data supplier berhasil di-import!")
                        st.rerun()
                with col_btn2:
                    if st.button("➕ Tambahkan ke Data Supplier Saat Ini", use_container_width=True):
                        st.session_state["df_supplier_state"] = pd.concat([st.session_state["df_supplier_state"], mapped_df], ignore_index=True)
                        st.success("Data supplier berhasil ditambahkan!")
                        st.rerun()

            except Exception as e:
                st.error(f"Gagal memproses file Excel: {e}")

    st.markdown("---")

    # --- TOMBOL AKSI & EXPORT DATA ---
    col1, col2 = st.columns([1, 2])
    with col1:
        if st.button("➕ Tambah Supplier Manual", type="primary", use_container_width=True):
            open_add_supplier_dialog()
            
    with col2:
        output = io.BytesIO()
        with pd.ExcelWriter(output) as writer:
            st.session_state["df_harga_bb"].to_excel(writer, index=False, sheet_name='UPDATE HARGA BB')
            st.session_state["df_supplier_state"].to_excel(writer, index=False, sheet_name='Supplier Link')
            st.session_state["df_dapur_state"].to_excel(writer, index=False, sheet_name='Data Dapur')
        
        st.download_button(
            label="📥 Download Seluruh Data System (Excel Multi-Sheet)",
            data=output.getvalue(),
            file_name="Master_Data_Procurement_System.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    st.markdown("---")
    
    # --- TAMPILAN DAFTAR SUPPLIER ---
    df_sup = st.session_state["df_supplier_state"].copy()
    st.markdown(f"##### 📋 **Total Supplier Terdaftar: {len(df_sup)}**")
    
    if not df_sup.empty:
        for idx, row in df_sup.iterrows():
            tkn = row.get("TOKEN", "")
            form_url = f"{WEB_APP_URL}?token={tkn}" if pd.notna(tkn) and tkn != "" else ""
            
            with st.expander(f"🤝 **{row.get('KODE SUPPLIER', '')} - {row.get('NAMA SUPPLIER', '')}** (WA: +{row.get('NO WA', '-')})"):
                c1, c2, c3 = st.columns([3, 1, 1])
                with c1:
                    st.write(f"**Supply BB:** {row.get('JENIS BB 1', '-')} | **PIC:** {row.get('PIC', '-')}")
                    st.write(f"🔗 **Form Link:** [{form_url}]({form_url})")
                with c2:
                    if st.button("✏️ Edit", key=f"btn_edit_sup_{idx}", use_container_width=True):
                        open_edit_supplier_dialog(idx)
                with c3:
                    if st.button("🗑️ Hapus", key=f"btn_del_sup_{idx}", type="secondary", use_container_width=True):
                        st.session_state["df_supplier_state"] = st.session_state["df_supplier_state"].drop(idx).reset_index(drop=True)
                        st.success("Supplier berhasil dihapus.")
                        st.rerun()
    else:
        st.info("Belum ada data supplier. Silakan upload file Excel di atas atau tambah secara manual.")

elif menu == "💬 WA & PO Generator":
    # ... (kode modul WA) ...

else:
    # ... (kode modul default) ...
