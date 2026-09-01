import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import random
import string

# ---------------------------------------------------------
# KONFIGURASI HALAMAN & UTILITY
# ---------------------------------------------------------
st.set_page_config(
    page_title="SPPG Procurement Engine | Koperasi YK",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

WEB_APP_URL = "https://sistem-procurement-koperasi-yk.streamlit.app"

def buat_token(length=32):
    return ''.join(random.choices(string.ascii_letters + string.digits, k=length))

# Custom CSS untuk Table Styling (Zebra Striping & Cards)
st.markdown("""
<style>
    .header-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        padding: 1.25rem 1.5rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    
    .table-header {
        background-color: #1e293b;
        color: white;
        padding: 10px;
        font-weight: bold;
        border-radius: 6px;
        margin-bottom: 8px;
    }
    
    .row-even {
        background-color: #f8fafc;
        padding: 8px 10px;
        border-radius: 6px;
        margin-bottom: 4px;
        border: 1px solid #e2e8f0;
        align-items: center;
    }
    
    .row-odd {
        background-color: #ffffff;
        padding: 8px 10px;
        border-radius: 6px;
        margin-bottom: 4px;
        border: 1px solid #e2e8f0;
        align-items: center;
    }

    .stSidebar [data-testid="stRadioButton"] > label {
        display: none;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# FUNCTION LOAD DATA EXCEL / XLSB
# ---------------------------------------------------------
@st.cache_data
def load_data():
    file_name = "Sistem Evaluasi Supplier & Dapur.xlsb"
    if not os.path.exists(file_name):
        file_name = "Sistem Evaluasi Supplier & Dapur.xlsx"
        
    if os.path.exists(file_name):
        try:
            xls = pd.ExcelFile(file_name)
            data_dict = {}
            for sheet in xls.sheet_names:
                df = pd.read_excel(file_name, sheet_name=sheet)
                data_dict[sheet] = df
            return data_dict, file_name
        except Exception as e:
            return None, str(e)
    return None, "File data tidak ditemukan."

data_excel, file_status = load_data()

# ---------------------------------------------------------
# INITIALIZE SESSION STATES
# ---------------------------------------------------------
# 1. State Dapur
if "df_dapur_state" not in st.session_state:
    if data_excel and isinstance(data_excel, dict) and "Data Dapur" in data_excel:
        raw_df = data_excel["Data Dapur"].copy()
        
        if "Unnamed" in str(raw_df.columns[0]):
            header_idx = raw_df[raw_df.apply(lambda row: row.astype(str).str.contains('NAMA DAPUR').any(), axis=1)].index
            if not header_idx.empty:
                idx = header_idx[0]
                raw_df.columns = raw_df.iloc[idx]
                raw_df = raw_df.iloc[idx+1:].reset_index(drop=True)
                raw_df = raw_df.dropna(how="all", subset=["NAMA DAPUR"])
        
        raw_df.columns = [str(c).strip().upper() for c in raw_df.columns]
        cols_to_keep = [c for c in ["KODE", "NAMA DAPUR", "ALAMAT", "PIC", "LATITUDE", "LONGTITUDE", "LONGITUDE", "KOTA/KABUPATEN"] if c in raw_df.columns]
        df_clean = raw_df[cols_to_keep].copy()
        
        if "LONGTITUDE" in df_clean.columns:
            df_clean = df_clean.rename(columns={"LONGTITUDE": "LONGITUDE"})
            
        df_clean["LATITUDE"] = pd.to_numeric(df_clean["LATITUDE"], errors="coerce")
        df_clean["LONGITUDE"] = pd.to_numeric(df_clean["LONGITUDE"], errors="coerce")
        
        st.session_state["df_dapur_state"] = df_clean.reset_index(drop=True)
    else:
        st.session_state["df_dapur_state"] = pd.DataFrame({
            "KODE": ["PAKEM", "NGPLK", "SLMN4"],
            "NAMA DAPUR": ["PAKEM (Hargobinangun)", "NGEMPLAK (Umbulmartani 1)", "SLEMAN 4 (Triharjo)"],
            "ALAMAT": ["Jl. Kaliurang Km 22", "Jl. Kaliurang Km 15.5", "Jl. Letkol Subadri"],
            "PIC": ["SHINTA", "SHINTA", "CAHYO"],
            "LATITUDE": [-7.618382, -7.675789, -7.700475],
            "LONGITUDE": [110.426078, 110.417100, 110.342646],
            "KOTA/KABUPATEN": ["SLEMAN", "SLEMAN", "SLEMAN"]
        })

# 2. State Supplier
if "df_supplier_state" not in st.session_state:
    st.session_state["df_supplier_state"] = pd.DataFrame(columns=[
        "KODE SUPPLIER", "NAMA SUPPLIER", "NO WA", 
        "JENIS BB 1", "JENIS BB 2", "JENIS BB 3", 
        "TOKEN", "LINK FORM", "PIC", "ALAMAT"
    ])

# 3. State Harga BB
if "df_harga_bb" not in st.session_state:
    st.session_state["df_harga_bb"] = pd.DataFrame(columns=[
        "KODE SUPPLIER", "NAMA SUPPLIER", "NAMA BB", "HARGA PER SATUAN", "SATUAN"
    ])

# ---------------------------------------------------------
# DIALOGS / POP-UP DAPUR
# ---------------------------------------------------------
@st.dialog("➕ Tambah Dapur SPPG Baru")
def open_dapur_dialog():
    st.write("Isi rincian dapur operasional baru di bawah ini:")
    with st.form("form_dapur_add"):
        c1, c2 = st.columns(2)
        with c1:
            kode = st.text_input("Kode Dapur (Contoh: BNTL1)")
            nama = st.text_input("Nama Dapur SPPG")
            pic = st.text_input("Nama PIC / Penanggung Jawab")
            kota = st.selectbox("Kota / Kabupaten", ["BANTUL", "SLEMAN", "GUNUNGKIDUL", "KULON PROGO", "YOGYAKARTA"])
        with c2:
            alamat = st.text_area("Alamat Lengkap Dapur", height=80)
            lat = st.number_input("Latitude", value=-7.7956, format="%.6f")
            lon = st.number_input("Longitude", value=110.3695, format="%.6f")
            
        if st.form_submit_button("✨ Simpan Dapur Baru", use_container_width=True):
            new_row = pd.DataFrame([{
                "KODE": kode,
                "NAMA DAPUR": nama,
                "ALAMAT": alamat,
                "PIC": pic,
                "LATITUDE": lat,
                "LONGITUDE": lon,
                "KOTA/KABUPATEN": kota
            }])
            st.session_state["df_dapur_state"] = pd.concat([st.session_state["df_dapur_state"], new_row], ignore_index=True)
            st.success("Dapur baru berhasil ditambahkan!")
            st.rerun()

@st.dialog("✏️ Edit Data Dapur SPPG")
def open_edit_dapur_dialog(index):
    df = st.session_state["df_dapur_state"]
    row = df.iloc[index]
    
    st.write(f"Mengubah data untuk **{row.get('NAMA DAPUR', 'Dapur')}**:")
    with st.form("form_dapur_edit"):
        c1, c2 = st.columns(2)
        with c1:
            kode = st.text_input("Kode Dapur", value=str(row.get("KODE", "")))
            nama = st.text_input("Nama Dapur SPPG", value=str(row.get("NAMA DAPUR", "")))
            pic = st.text_input("Nama PIC", value=str(row.get("PIC", "")))
            
            list_kota = ["BANTUL", "SLEMAN", "GUNUNGKIDUL", "KULON PROGO", "YOGYAKARTA"]
            curr_kota = str(row.get("KOTA/KABUPATEN", "")).upper()
            idx_kota = list_kota.index(curr_kota) if curr_kota in list_kota else 0
            kota = st.selectbox("Kota / Kabupaten", list_kota, index=idx_kota)
        with c2:
            alamat = st.text_area("Alamat Lengkap Dapur", value=str(row.get("ALAMAT", "")), height=80)
            lat = st.number_input("Latitude", value=float(row.get("LATITUDE", -7.7956) if pd.notna(row.get("LATITUDE")) else -7.7956), format="%.6f")
            lon = st.number_input("Longitude", value=float(row.get("LONGITUDE", 110.3695) if pd.notna(row.get("LONGITUDE")) else 110.3695), format="%.6f")
            
        if st.form_submit_button("💾 Update Perubahan Data", use_container_width=True):
            st.session_state["df_dapur_state"].at[index, "KODE"] = kode
            st.session_state["df_dapur_state"].at[index, "NAMA DAPUR"] = nama
            st.session_state["df_dapur_state"].at[index, "ALAMAT"] = alamat
            st.session_state["df_dapur_state"].at[index, "PIC"] = pic
            st.session_state["df_dapur_state"].at[index, "LATITUDE"] = lat
            st.session_state["df_dapur_state"].at[index, "LONGITUDE"] = lon
            st.session_state["df_dapur_state"].at[index, "KOTA/KABUPATEN"] = kota
            st.success("Data dapur berhasil diperbarui!")
            st.rerun()

@st.dialog("🗑️ Konfirmasi Hapus Data Dapur")
def open_delete_dapur_dialog(index):
    df = st.session_state["df_dapur_state"]
    nama_dapur = df.iloc[index].get("NAMA DAPUR", "Dapur ini")
    
    st.warning(f"Apakah Anda yakin ingin menghapus data **{nama_dapur}**?")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("❌ Batal", use_container_width=True):
            st.rerun()
    with col2:
        if st.button("🗑️ Ya, Hapus Data", type="primary", use_container_width=True):
            st.session_state["df_dapur_state"] = st.session_state["df_dapur_state"].drop(index).reset_index(drop=True)
            st.success("Data dapur berhasil dihapus!")
            st.rerun()

# ---------------------------------------------------------
# SIDEBAR NAVIGASI
# ---------------------------------------------------------
with st.sidebar:
    st.markdown("### ⚡ **SPPG Engine**")
    st.caption("Koperasi YK")
    st.markdown("---")
    
    st.subheader("🚀 Modul Operasional")
    
    modul_options = [
        "📊 Dashboard & HET",
        "🏬 Kelola Data Dapur",
        "🤝 Data Supplier & Link Form",
        "💬 WA & PO Generator",
        "🚛 Matriks Jarak",
        "🎯 Scoring & Evaluasi",
        "📈 HET & Komparasi Pasar"
    ]
    
    menu = st.radio("Pilih Modul:", modul_options, index=1)
    st.markdown("---")
    st.caption(f"Status Data: 🟢 {file_status}")

# ---------------------------------------------------------
# HEADER UTAMA
# ---------------------------------------------------------
st.markdown(f"""
<div class="header-card">
    <div style="display: flex; justify-content: space-between; align-items: center;">
        <div>
            <h2 style="margin:0; color: white;">🏭 Enterprise Procurement System</h2>
            <p style="margin:0; opacity: 0.8;">Koperasi YK — Sistem Evaluasi Supplier & Dapur SPPG</p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)


# ---------------------------------------------------------
# MODUL 1: DASHBOARD & HET
# ---------------------------------------------------------
if menu == "📊 Dashboard & HET":
    st.subheader("📊 Dashboard Utama & Pemantauan HET")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Dapur SPPG", len(st.session_state["df_dapur_state"]), delta="Aktif")
    m2.metric("Total Supplier", len(st.session_state["df_supplier_state"]), delta="Terdaftar")
    m3.metric("Kategori Barang", "6 Kategori", delta="Lengkap")
    m4.metric("Rata-rata Ketepatan", "94.2%", delta="+1.5%")


# ---------------------------------------------------------
# MODUL 2: KELOLA DATA DAPUR SPPG
# ---------------------------------------------------------
elif menu == "🏬 Kelola Data Dapur":
    st.subheader("🏬 Manajemen Data Dapur SPPG & Peta Sebaran")
    
    df_dapur = st.session_state["df_dapur_state"]
    
    col_tambah, _ = st.columns([1, 3])
    with col_tambah:
        if st.button("➕ Tambah Dapur Baru", type="primary", use_container_width=True):
            open_dapur_dialog()
            
    st.markdown("<br>", unsafe_allow_html=True)
    
    # --- PETA SEBARAN DAPUR SPPG ---
    st.markdown("##### 🗺️ **Peta Sebaran Lokasi Dapur SPPG**")
    map_data = df_dapur.dropna(subset=["LATITUDE", "LONGITUDE"]).copy()
    
    if not map_data.empty:
        map_df = pd.DataFrame({
            "lat": map_data["LATITUDE"].astype(float),
            "lon": map_data["LONGITUDE"].astype(float)
        })
        st.map(map_df, zoom=10, use_container_width=True)
    else:
        st.warning("Data Koordinat (Latitude/Longitude) belum terisi dengan benar.")
        
    st.markdown("---")
    
    # --- TABEL DATA DAPUR SELANG-SELING (ZEBRA STRIPING) ---
    st.markdown("##### 📋 **Daftar Dapur SPPG Operasional**")
    
    if not df_dapur.empty:
        h_cols = st.columns([0.8, 1.8, 2.5, 1.0, 1.1, 1.1, 0.9, 0.5, 0.5])
        headers = ["KODE", "NAMA DAPUR", "ALAMAT", "PIC", "LATITUDE", "LONGITUDE", "MAPS", "EDIT", "HAPUS"]
        
        for col, h in zip(h_cols, headers):
            col.markdown(f"**{h}**")
        st.markdown("<hr style='margin-top:0; margin-bottom:10px;'>", unsafe_allow_html=True)
        
        for i, r in df_dapur.iterrows():
            bg_color = "#f8fafc" if i % 2 == 0 else "#ffffff"
            
            with st.container():
                st.markdown(
                    f"""
                    <div style="background-color: {bg_color}; padding: 6px 10px; border-radius: 8px; border: 1px solid #e2e8f0; margin-bottom: 6px;">
                    """,
                    unsafe_allow_html=True
                )
                
                c_kode, c_nama, c_alamat, c_pic, c_lat, c_lon, c_map, c_edit, c_del = st.columns([0.8, 1.8, 2.5, 1.0, 1.1, 1.1, 0.9, 0.5, 0.5])
                
                lat_val = r.get("LATITUDE", None)
                lon_val = r.get("LONGITUDE", None)
                
                c_kode.write(f"**{r.get('KODE', '-')}**")
                c_nama.write(r.get("NAMA DAPUR", "-"))
                c_alamat.write(r.get("ALAMAT", "-"))
                c_pic.write(r.get("PIC", "-"))
                c_lat.write(f"{lat_val:.5f}" if pd.notna(lat_val) else "-")
                c_lon.write(f"{lon_val:.5f}" if pd.notna(lon_val) else "-")
                
                if pd.notna(lat_val) and pd.notna(lon_val):
                    gmap_url = f"https://www.google.com/maps/search/?api=1&query={lat_val},{lon_val}"
                    c_map.markdown(f"[📍 Buka]({gmap_url})")
                else:
                    c_map.write("-")
                
                if c_edit.button("✏️", key=f"edit_{i}"):
                    open_edit_dapur_dialog(i)
                    
                if c_del.button("🗑️", key=f"del_{i}"):
                    open_delete_dapur_dialog(i)
                
                st.markdown("</div>", unsafe_allow_html=True)
    else:
        st.info("Belum ada data dapur. Klik 'Tambah Dapur Baru' untuk menambahkan.")


# ---------------------------------------------------------
# MODUL 3: KELOLA DATA SUPPLIER & LINK WEB APP
# ---------------------------------------------------------
elif menu == "🤝 Data Supplier & Link Form":
    st.subheader("🤝 Kelola Data Supplier & Link Web App")
    
    # --- FITUR UPLOAD FILE EXCEL / SPREADSHEET ---
    with st.expander("📂 **Upload Data Supplier dari File Excel / Google Sheets**", expanded=True):
        st.write("Unggah file `.xlsx`, `.xls`, atau `.csv` yang berisi daftar supplier kamu.")
        uploaded_file = st.file_uploader("Pilih file Excel Supplier:", type=["xlsx", "xls", "xlsb", "csv"])
        
        if uploaded_file is not None:
            try:
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
                
                # Mapping Kolom dari Spreadsheet ke Format System
                mapped_df = pd.DataFrame()
                
                col_code = next((c for c in df_upload.columns if 'SUPP CODE' in c.upper() or 'KODE' in c.upper()), None)
                mapped_df["KODE SUPPLIER"] = df_upload[col_code].astype(str) if col_code else [f"SUP-{i+1:03d}" for i in range(len(df_upload))]
                
                col_name = next((c for c in df_upload.columns if 'SUPPLIER NAME' in c.upper() or 'NAMA' in c.upper()), None)
                mapped_df["NAMA SUPPLIER"] = df_upload[col_name].astype(str) if col_name else "Tanpa Nama"
                
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

                col_bb = next((c for c in df_upload.columns if 'SUPPLY BB' in c.upper() or 'JENIS' in c.upper() or 'KATEGORI' in c.upper()), None)
                mapped_df["JENIS BB 1"] = df_upload[col_bb].astype(str) if col_bb else "-"
                mapped_df["JENIS BB 2"] = "-"
                mapped_df["JENIS BB 3"] = "-"

                mapped_df["TOKEN"] = [buat_token() for _ in range(len(mapped_df))]
                mapped_df["LINK FORM"] = mapped_df["TOKEN"].apply(lambda t: f"{WEB_APP_URL}?token={t}")

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

    # --- TOMBOL EXPORT SELURUH DATA ---
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
                c1, c2 = st.columns([4, 1])
                with c1:
                    st.write(f"**Supply BB:** {row.get('JENIS BB 1', '-')} | **PIC:** {row.get('PIC', '-')}")
                    st.write(f"🔗 **Form Link:** [{form_url}]({form_url})")
                with c2:
                    if st.button("🗑️ Hapus", key=f"btn_del_sup_{idx}", type="secondary", use_container_width=True):
                        st.session_state["df_supplier_state"] = st.session_state["df_supplier_state"].drop(idx).reset_index(drop=True)
                        st.success("Supplier berhasil dihapus.")
                        st.rerun()
    else:
        st.info("Belum ada data supplier. Silakan upload file Excel di atas.")


# ---------------------------------------------------------
# MODUL 4: WA & PO GENERATOR (DENGAN EDIT KATEGORI SUPPLIER)
# ---------------------------------------------------------
elif menu == "💬 WA & PO Generator":
    st.subheader("💬 WA Message & PO Link Generator")
    st.caption("Kelola kategori barang yang disuplai oleh masing-masing vendor dan buat draft pesan WhatsApp beserta link penawaran harganya.")

    df_sup = st.session_state["df_supplier_state"]

    if df_sup.empty:
        st.warning("⚠️ Data supplier masih kosong. Silakan isi atau import data di menu '🤝 Data Supplier & Link Form' terlebih dahulu.")
    else:
        # Master list kategori sesuai gambar spreadsheet
        LIST_KATEGORI = [
            "Ayam", "Beras", "Buah", "Cookies", "Daging", "Ikan", 
            "Keju", "Olahan", "Sayur", "Sembako", "Susu", "Tahu", 
            "Tempe", "Telur Ayam", "Telur Bebek", "Telur Puyuh"
        ]

        # 1. PILIH SUPPLIER
        st.markdown("##### 1️⃣ **Pilih Supplier & Atur Kategori Barang**")
        
        list_supplier_display = [f"{r['KODE SUPPLIER']} - {r['NAMA SUPPLIER']}" for _, r in df_sup.iterrows()]
        selected_sup_str = st.selectbox("Pilih Supplier:", list_supplier_display)
        
        # Ambil index data supplier terpilih
        idx_sup = list_supplier_display.index(selected_sup_str)
        row_sup = df_sup.iloc[idx_sup]

        # 2. EDIT KATEGORI SUPPLIER (JENIS BB 1 s.d JENIS BB 6)
        st.markdown("---")
        st.write(f"Edit kategori bahan baku yang disuplai oleh **{row_sup['NAMA SUPPLIER']}**:")

        # Mengambil data kategori eksisting
        kat_existing = []
        for col_bb in ["JENIS BB 1", "JENIS BB 2", "JENIS BB 3", "JENIS BB 4", "JENIS BB 5", "JENIS BB 6"]:
            val = row_sup.get(col_bb, "-")
            if pd.notna(val) and str(val).strip() not in ["-", "", "nan", "None"]:
                # Match casing sesuai LIST_KATEGORI
                matched = next((k for k in LIST_KATEGORI if k.upper() == str(val).strip().upper()), None)
                if matched:
                    kat_existing.append(matched)

        # UI Form Edit Kategori (Multiselect maks 6)
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
                # Update ke session state
                for i in range(1, 7):
                    col_name = f"JENIS BB {i}"
                    val_to_save = selected_kategori[i-1] if i <= len(selected_kategori) else "-"
                    st.session_state["df_supplier_state"].at[idx_sup, col_name] = val_to_save
                
                st.success(f"Kategori untuk {row_sup['NAMA SUPPLIER']} berhasil diperbarui!")
                st.rerun()

        # 3. GENERATE LINK & PESAN WA
        st.markdown("---")
        st.markdown("##### 2️⃣ **Preview Link & Pesan WhatsApp**")

        # Format URL dengan Token & Parameter Kategori
        token = row_sup.get("TOKEN", "")
        kat_param = ",".join([k.replace(" ", "_") for k in selected_kategori]) if selected_kategori else "ALL"
        
        # Link Penawaran khusus kategori supplier
        link_form_khusus = f"{WEB_APP_URL}?token={token}&kat={kat_param}"

        # Template Pesan WhatsApp
        no_wa = str(row_sup.get("NO WA", "")).replace("+", "").replace(" ", "").replace("-", "")
        nama_supplier = row_sup.get("NAMA SUPPLIER", "Bapak/Ibu Vendor")
        kategori_txt = ", ".join(selected_kategori) if selected_kategori else "Bahan Baku"

        pesan_wa = (
            f"Halo *{nama_supplier}*,\n\n"
            f"Kami dari tim Procurement Koperasi YK/SPPG. "
            f"Mohon untuk mengisi update harga penawaran harian/mingguan untuk kategori *[{kategori_txt}]* "
            f"melalui link form resmi berikut ini:\n\n"
            f"🔗 {link_form_khusus}\n\n"
            f"Terima kasih atas kerja samanya."
        )

        # Display Hasil Generator
        st.text_area("Draft Pesan WhatsApp:", value=pesan_wa, height=180)

        col_wa1, col_wa2 = st.columns([1, 1])
        with col_wa1:
            # Tombol Direct Open ke WA Web / App
            import urllib.parse
            encoded_msg = urllib.parse.quote(pesan_wa)
            wa_url = f"https://wa.me/{no_wa}?text={encoded_msg}"
            
            st.markdown(
                f'<a href="{wa_url}" target="_blank" style="text-decoration:none;">'
                f'<div style="background-color:#25D366; color:white; padding:10px; border-radius:8px; text-align:center; font-weight:bold;">'
                f'📲 Kirim via WhatsApp (+{no_wa})</div></a>',
                unsafe_allow_html=True
            )
            
        with col_wa2:
            st.info(f"💡 Link ini secara otomatis membatasi Form Supplier agar **hanya menampilkan barang kategori**: `{kategori_txt}`.")
