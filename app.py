import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import random
import string
import urllib.parse

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
DATA_HARGA_FILE = "Data_Penawaran_Supplier_Master.csv"  # File penyimpan permanen data penawaran harga

def buat_token(length=32):
    return ''.join(random.choices(string.ascii_letters + string.digits, k=length))

def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0
    lat1, lon1, lat2, lon2 = map(np.radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = np.sin(dlat / 2.0)**2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon / 2.0)**2
    c = 2 * np.arcsin(np.sqrt(a))
    return R * c

# Custom CSS untuk Table Styling & Header
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
    }
    .row-odd {
        background-color: #ffffff;
        padding: 8px 10px;
        border-radius: 6px;
        margin-bottom: 4px;
        border: 1px solid #e2e8f0;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# FUNCTION LOAD & SAVE DATA PERMANEN
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

# Helper Penyimpanan Permanen Data Penawaran Supplier
def load_penawaran_permanen():
    if os.path.exists(DATA_HARGA_FILE):
        return pd.read_csv(DATA_HARGA_FILE)
    return pd.DataFrame(columns=["KODE SUPPLIER", "NAMA SUPPLIER", "P/N", "JENIS BB", "ITEM BB", "SATUAN", "HARGA PENAWARAN", "CATATAN"])

def save_penawaran_permanen(df_new_data):
    df_existing = load_penawaran_permanen()
    df_combined = pd.concat([df_existing, df_new_data], ignore_index=True)
    df_combined = df_combined.drop_duplicates(subset=["KODE SUPPLIER", "ITEM BB"], keep="last")
    df_combined.to_csv(DATA_HARGA_FILE, index=False)
    return df_combined

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
    st.session_state["df_supplier_state"] = pd.DataFrame({
        "KODE SUPPLIER": ["SUP-001", "SUP-002"],
        "NAMA SUPPLIER": ["CV Sumber Pangan", "UD Tani Makmur"],
        "NO WA": ["6281234567890", "6289876543210"],
        "JENIS BB 1": ["Beras", "Sayur"],
        "JENIS BB 2": ["Sembako", "Buah"],
        "JENIS BB 3": ["-", "-"],
        "TOKEN": [buat_token(), buat_token()],
        "LINK FORM": [f"{WEB_APP_URL}?token=12345", f"{WEB_APP_URL}?token=67890"],
        "PIC": ["Budi", "Siti"],
        "ALAMAT": ["Jl. Solo Km 9", "Jl. Magelang Km 5"],
        "LATITUDE": [-7.780000, -7.750000],
        "LONGITUDE": [110.400000, 110.360000],
        "RATING": [4.8, 4.5],
        "KETEPATAN_KIRIM": [95, 90]
    })

# 3. State Master Bahan Baku (525 Master Item)
if "df_bb_state" not in st.session_state:
    if data_excel and isinstance(data_excel, dict) and "Master Bahan Baku" in data_excel:
        st.session_state["df_bb_state"] = data_excel["Master Bahan Baku"].copy()
    else:
        st.session_state["df_bb_state"] = pd.DataFrame({
            "P/N": [f"BB-{i+1:03d}" for i in range(10)],
            "JENIS BB": ["Ayam", "Beras", "Sayur", "Sembako", "Telur Ayam", "Daging", "Ikan", "Bumbu", "Susu", "Buah"],
            "ITEM BB": ["Ayam Broiler Utuh", "Beras C4 Super", "Bayam Fresh", "Minyak Goreng 2L", "Telur Ayam Ras", "Daging Sapi Dadu", "Ikan Gurame", "Bawang Merah", "Susu UHT 1L", "Pisang Ambon"],
            "SATUAN": ["Kg", "Kg", "Ikat", "Pcs", "Kg", "Kg", "Kg", "Kg", "Karton", "Sisir"],
            "HET": [38000, 15000, 3000, 34000, 28000, 130000, 40000, 35000, 180000, 20000]
        })

if "HET" not in st.session_state["df_bb_state"].columns:
    st.session_state["df_bb_state"]["HET"] = 30000

# 4. State Harga BB / Penawaran Supplier (Di-load dari CSV Permanen)
st.session_state["df_harga_bb"] = load_penawaran_permanen()

# ---------------------------------------------------------
# MODE VENDOR EXTERNAL (PENGISIAN SPREADSHEET VIA TOKEN LINK)
# ---------------------------------------------------------
query_params = st.query_params
if "token" in query_params:
    token_diterima = query_params["token"]
    kategori_diterima = query_params.get("kat", "ALL")
    
    df_sup = st.session_state.get("df_supplier_state", pd.DataFrame())
    supplier_info = None
    
    if not df_sup.empty and "TOKEN" in df_sup.columns:
        match = df_sup[df_sup["TOKEN"] == token_diterima]
        if not match.empty:
            supplier_info = match.iloc[0]

    st.title("📝 Portal Penawaran Harga Bahan Baku")
    st.caption("Koperasi YK — Sistem Informasi Procurement & SPPG")
    st.markdown("---")

    if supplier_info is not None:
        st.success(f"Selamat datang, **{supplier_info['NAMA SUPPLIER']}** (Kode: `{supplier_info['KODE SUPPLIER']}`)!")
        kode_sup = supplier_info['KODE SUPPLIER']
        nama_sup = supplier_info['NAMA SUPPLIER']
    else:
        st.info("Form Penawaran Harga Barang Operasional — Mode Akses Vendor External")
        kode_sup = "GUEST"
        nama_sup = "Supplier External"

    tab_bulk, tab_single = st.tabs(["📊 Update Spreadsheet Massal (Rekomendasi)", "✏️ Input Manual Satuan"])

    with tab_bulk:
        st.markdown("##### 💡 Petunjuk Pengisian Massal:")
        st.caption("1. Gunakan filter pencarian/kategori di bawah untuk menemukan bahan baku yang Anda suplai.\n2. Klik dua kali pada kolom **HARGA PENAWARAN (RP)** untuk mengetikkan harga.\n3. Klik tombol **Simpan Semua Penawaran** di bagian bawah setelah selesai.")
        
        df_master_bb = st.session_state["df_bb_state"].copy()
        df_existing_harga = st.session_state["df_harga_bb"]
        df_existing_sup = df_existing_harga[df_existing_harga["KODE SUPPLIER"] == kode_sup] if not df_existing_harga.empty else pd.DataFrame()

        if not df_existing_sup.empty:
            df_merged = pd.merge(df_master_bb, df_existing_sup[["ITEM BB", "HARGA PENAWARAN", "CATATAN"]], on="ITEM BB", how="left")
        else:
            df_merged = df_master_bb.copy()
            df_merged["HARGA PENAWARAN"] = 0
            df_merged["CATATAN"] = ""

        df_merged["HARGA PENAWARAN"] = df_merged["HARGA PENAWARAN"].fillna(0).astype(int)
        df_merged["CATATAN"] = df_merged["CATATAN"].fillna("")

        col_f1, col_f2 = st.columns([1, 2])
        with col_f1:
            kat_list = ["Semua Kategori"] + list(df_merged["JENIS BB"].dropna().unique())
            sel_kat = st.selectbox("Filter Kategori Barang:", kat_list)
        with col_f2:
            search_item = st.text_input("🔍 Cari Nama Bahan Baku:", placeholder="Contoh: Telur, Minyak, Ayam...")

        df_disp = df_merged.copy()
        if sel_kat != "Semua Kategori":
            df_disp = df_disp[df_disp["JENIS BB"] == sel_kat]
        if search_item:
            df_disp = df_disp[df_disp["ITEM BB"].str.contains(search_item, case=False, na=False)]

        edited_df = st.data_editor(
            df_disp[["P/N", "JENIS BB", "ITEM BB", "SATUAN", "HARGA PENAWARAN", "CATATAN"]],
            column_config={
                "P/N": st.column_config.TextColumn("Kode P/N", disabled=True),
                "JENIS BB": st.column_config.TextColumn("Kategori", disabled=True),
                "ITEM BB": st.column_config.TextColumn("Nama Bahan Baku", disabled=True),
                "SATUAN": st.column_config.TextColumn("Satuan", disabled=True),
                "HARGA PENAWARAN": st.column_config.NumberColumn(
                    "Harga Penawaran (Rp)",
                    help="Masukkan nominal harga tanpa titik",
                    min_value=0,
                    step=500,
                    format="Rp %d"
                ),
                "CATATAN": st.column_config.TextColumn("Catatan / Merek (Opsional)")
            },
            disabled=["P/N", "JENIS BB", "ITEM BB", "SATUAN"],
            hide_index=True,
            use_container_width=True,
            num_rows="fixed",
            height=450
        )

        if st.button("🚀 SIMPAN SEMUA PENAWARAN HARGA", type="primary", use_container_width=True):
            df_valid = edited_df[edited_df["HARGA PENAWARAN"] > 0].copy()
            if df_valid.empty:
                st.error("Mohon isi minimal satu harga penawaran di atas Rp 0.")
            else:
                df_valid["KODE SUPPLIER"] = kode_sup
                df_valid["NAMA SUPPLIER"] = nama_sup
                cols_order = ["KODE SUPPLIER", "NAMA SUPPLIER", "P/N", "JENIS BB", "ITEM BB", "SATUAN", "HARGA PENAWARAN", "CATATAN"]
                df_final = df_valid[cols_order]
                
                save_penawaran_permanen(df_final)
                st.session_state["df_harga_bb"] = load_penawaran_permanen()
                st.balloons()
                st.success(f"Berhasil! {len(df_final)} penawaran harga dari **{nama_sup}** telah tersimpan di sistem.")

    with tab_single:
        with st.form("form_input_harga_supplier_single"):
            st.subheader("🛒 Input Single Item Penawaran")
            nama_barang = st.text_input("Nama Bahan Baku / Barang", placeholder="Contoh: Telur Ayam Ras / Minyak Goreng")
            
            c1, c2 = st.columns(2)
            with c1:
                harga_penawaran = st.number_input("Harga Penawaran (Rp)", min_value=0, step=500, value=10000)
                satuan = st.selectbox("Satuan", ["Kg", "Liter", "Ikat", "Pcs", "Karton", "Ekor", "Pack", "Sisir"])
            with c2:
                kategori_pilihan = st.selectbox(
                    "Kategori Barang", 
                    options=[
                        "Ayam", "Beras", "Buah", "Cookies", "Daging", "Ikan", 
                        "Keju", "Olahan", "Sayur", "Sembako", "Susu", "Tahu", 
                        "Tempe", "Telur Ayam", "Telur Bebek", "Telur Puyuh"
                    ]
                )
                catatan = st.text_input("Catatan / Merek (Opsional)", placeholder="Contoh: Ukuran Medium, Fresh")

            submit_harga = st.form_submit_button("🚀 Kirim Penawaran Item Ini", use_container_width=True)

            if submit_harga:
                if nama_barang.strip() == "":
                    st.error("Mohon isi nama barang terlebih dahulu.")
                else:
                    new_entry = pd.DataFrame([{
                        "KODE SUPPLIER": kode_sup,
                        "NAMA SUPPLIER": nama_sup,
                        "P/N": "CUSTOM",
                        "JENIS BB": kategori_pilihan,
                        "ITEM BB": nama_barang,
                        "SATUAN": satuan,
                        "HARGA PENAWARAN": harga_penawaran,
                        "CATATAN": catatan
                    }])
                    save_penawaran_permanen(new_entry)
                    st.session_state["df_harga_bb"] = load_penawaran_permanen()
                    st.balloons()
                    st.success(f"Berhasil! Penawaran untuk **{nama_barang}** seharga **Rp {harga_penawaran:,}** telah tersimpan.")

    st.stop()

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
                "KODE": kode, "NAMA DAPUR": nama, "ALAMAT": alamat,
                "PIC": pic, "LATITUDE": lat, "LONGITUDE": lon, "KOTA/KABUPATEN": kota
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
# SIDEBAR NAVIGASI ADMIN
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
    menu = st.radio("Pilih Modul:", modul_options, index=0)
    st.markdown("---")
    st.caption(f"Status Data: 🟢 {file_status}")

# HEADER UTAMA
st.markdown("""
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
    m3.metric("Penawaran Masuk", len(st.session_state["df_harga_bb"]), delta="Item")
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
    st.markdown("##### 🗺️ **Peta Sebaran Lokasi Dapur SPPG**")
    map_data = df_dapur.dropna(subset=["LATITUDE", "LONGITUDE"]).copy()
    
    if not map_data.empty:
        map_df = pd.DataFrame({
            "lat": map_data["LATITUDE"].astype(float),
            "lon": map_data["LONGITUDE"].astype(float)
        })
        st.map(map_df, zoom=10, use_container_width=True)
    else:
        st.warning("Data Koordinat belum terisi dengan benar.")
        
    st.markdown("---")
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
                st.markdown(f'<div style="background-color: {bg_color}; padding: 6px 10px; border-radius: 8px; border: 1px solid #e2e8f0; margin-bottom: 6px;">', unsafe_allow_html=True)
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

# ---------------------------------------------------------
# MODUL 3: DATA SUPPLIER & EXPORT/IMPORT
# ---------------------------------------------------------
elif menu == "🤝 Data Supplier & Link Form":
    st.subheader("🤝 Kelola Data Supplier & Link Web App")
    
    with st.expander("📂 **Upload Data Supplier dari File Excel / Google Sheets**", expanded=True):
        uploaded_file = st.file_uploader("Pilih file Excel Supplier:", type=["xlsx", "xls", "xlsb", "csv"])
        
        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith('.csv'):
                    df_upload = pd.read_csv(uploaded_file)
                else:
                    df_upload = pd.read_excel(uploaded_file)
                
                df_upload.columns = [str(c).strip() for c in df_upload.columns]
                mapped_df = pd.DataFrame()
                
                col_code = next((c for c in df_upload.columns if 'SUPP CODE' in c.upper() or 'KODE' in c.upper()), None)
                mapped_df["KODE SUPPLIER"] = df_upload[col_code].astype(str) if col_code else [f"SUP-{i+1:03d}" for i in range(len(df_upload))]
                
                col_name = next((c for c in df_upload.columns if 'SUPPLIER NAME' in c.upper() or 'NAMA' in c.upper()), None)
                mapped_df["NAMA SUPPLIER"] = df_upload[col_name].astype(str) if col_name else "Tanpa Nama"
                
                col_phone = next((c for c in df_upload.columns if 'PHONE' in c.upper() or 'WA' in c.upper() or 'TELP' in c.upper()), None)
                if col_phone:
                    def clean_phone(p):
                        p_str = str(p).replace("-", "").replace(" ", "").replace("+", "").split(".")[0].strip()
                        if p_str.startswith("0"): return "62" + p_str[1:]
                        elif not p_str.startswith("62") and len(p_str) > 5: return "62" + p_str
                        return p_str
                    mapped_df["NO WA"] = df_upload[col_phone].apply(clean_phone)
                else:
                    mapped_df["NO WA"] = "6285169796974"

                mapped_df["JENIS BB 1"] = "-"
                mapped_df["TOKEN"] = [buat_token() for _ in range(len(mapped_df))]
                mapped_df["LINK FORM"] = mapped_df["TOKEN"].apply(lambda t: f"{WEB_APP_URL}?token={t}")

                st.session_state["df_supplier_state"] = mapped_df
                st.success(f"✅ Berhasil memuat {len(mapped_df)} data supplier!")
                st.rerun()
            except Exception as e:
                st.error(f"Gagal memproses file: {e}")

    st.markdown("---")
    st.dataframe(st.session_state["df_supplier_state"], use_container_width=True)

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

# ---------------------------------------------------------
# MODUL 4: WA & PO GENERATOR
# ---------------------------------------------------------
elif menu == "💬 WA & PO Generator":
    st.subheader("💬 WhatsApp Generator - Update Penawaran Harga Supplier")
    df_supplier = st.session_state["df_supplier_state"].copy()

    if df_supplier.empty:
        st.warning("⚠️ Data supplier masih kosong. Silakan upload data supplier terlebih dahulu.")
    else:
        st.markdown("### 📝 Template Pesan WhatsApp")
        template_default = """Halo Bapak/Ibu {nama_supplier} 👋

Kami dari *Koperasi YK - Procurement SPPG*.

Mohon bantuan untuk melakukan update penawaran harga bahan baku yang Bapak/Ibu supply.

Silakan isi melalui link berikut:

{link_form}

Mohon harga dapat diperbarui sesuai harga terbaru.

Terima kasih atas kerja samanya 🙏

*Koperasi YK*
Sistem Procurement SPPG"""

        template_pesan = st.text_area(
            "Edit Template Pesan",
            value=template_default,
            height=280,
            help="Gunakan {nama_supplier} untuk nama supplier dan {link_form} untuk link penawaran."
        )

        st.markdown("---")
        st.markdown("### 🎯 Pilih Supplier")
        col1, col2 = st.columns([2, 1])

        with col1:
            list_supplier = df_supplier["NAMA SUPPLIER"].dropna().tolist()
            selected_supplier = st.multiselect(
                "Pilih Supplier yang akan dikirim",
                options=list_supplier,
                default=list_supplier
            )

        with col2:
            st.metric("Supplier Dipilih", len(selected_supplier))

        st.markdown("---")
        st.markdown("### 📱 Kirim Link Penawaran Harga")

        df_selected = df_supplier[df_supplier["NAMA SUPPLIER"].isin(selected_supplier)].copy()

        if df_selected.empty:
            st.info("Pilih minimal satu supplier.")
        else:
            h1, h2, h3, h4, h5 = st.columns([0.5, 2.5, 1.5, 3, 1])
            h1.markdown("**NO**")
            h2.markdown("**SUPPLIER**")
            h3.markdown("**NO WA**")
            h4.markdown("**LINK FORM**")
            h5.markdown("**KIRIM**")
            st.markdown("---")

            for idx, row in df_selected.iterrows():
                nomor = row.get("NO WA", "")
                nama = row.get("NAMA SUPPLIER", "")
                link = row.get("LINK FORM", "")

                pesan = template_pesan.format(nama_supplier=nama, link_form=link)
                pesan_encoded = urllib.parse.quote(pesan)
                wa_link = f"https://wa.me/{nomor}?text={pesan_encoded}"

                c1, c2, c3, c4, c5 = st.columns([0.5, 2.5, 1.5, 3, 1])
                c1.write(idx + 1)
                c2.write(nama)
                c3.write(nomor)
                c4.code(link, language=None)
                c5.link_button("💬 WA", wa_link, use_container_width=True)

        st.markdown("---")
        st.markdown("### 👀 Preview Pesan")
        if not df_selected.empty:
            preview_supplier = st.selectbox("Pilih supplier untuk preview", df_selected["NAMA SUPPLIER"].tolist())
            preview_row = df_selected[df_selected["NAMA SUPPLIER"] == preview_supplier].iloc[0]
            preview_message = template_pesan.format(
                nama_supplier=preview_row["NAMA SUPPLIER"],
                link_form=preview_row["LINK FORM"]
            )
            st.info(preview_message)

# ---------------------------------------------------------
# MODUL 5: MATRIKS JARAK
# ---------------------------------------------------------
elif menu == "🚛 Matriks Jarak":
    st.subheader("🚛 Matriks Jarak Supplier ke Dapur SPPG")
    df_sup = st.session_state["df_supplier_state"]
    df_dapur = st.session_state["df_dapur_state"]

    if df_sup.empty or df_dapur.empty:
        st.warning("Data Supplier atau Data Dapur masih kosong. Silakan lengkapi data terlebih dahulu.")
    else:
        cost_per_km = st.number_input("Estimasi Biaya Logistik per Km (Rp)", value=2500, step=500)
        
        results = []
        for _, sup in df_sup.iterrows():
            lat_s = sup.get("LATITUDE", np.nan)
            lon_s = sup.get("LONGITUDE", np.nan)
            
            for _, dap in df_dapur.iterrows():
                lat_d = dap.get("LATITUDE", np.nan)
                lon_d = dap.get("LONGITUDE", np.nan)
                
                if pd.notna(lat_s) and pd.notna(lon_s) and pd.notna(lat_d) and pd.notna(lon_d):
                    dist = haversine(lat_s, lon_s, lat_d, lon_d)
                else:
                    dist = np.nan
                    
                est_cost = dist * cost_per_km if pd.notna(dist) else np.nan
                
                results.append({
                    "KODE SUPPLIER": sup.get("KODE SUPPLIER", "-"),
                    "NAMA SUPPLIER": sup.get("NAMA SUPPLIER", "-"),
                    "KODE DAPUR": dap.get("KODE", "-"),
                    "NAMA DAPUR": dap.get("NAMA DAPUR", "-"),
                    "JARAK (KM)": round(dist, 2) if pd.notna(dist) else "-",
                    "ESTIMASI BIAYA LOGISTIK": f"Rp {int(est_cost):,}" if pd.notna(est_cost) else "-"
                })
        
        df_matriks = pd.DataFrame(results)
        
        st.markdown("##### 🔍 Filter Matriks Jarak")
        f_col1, f_col2 = st.columns(2)
        with f_col1:
            sel_dapur_filter = st.selectbox("Filter Dapur:", ["Semua Dapur"] + list(df_dapur["NAMA DAPUR"].unique()))
        with f_col2:
            sel_sup_filter = st.selectbox("Filter Supplier:", ["Semua Supplier"] + list(df_sup["NAMA SUPPLIER"].unique()))
            
        df_disp_matriks = df_matriks.copy()
        if sel_dapur_filter != "Semua Dapur":
            df_disp_matriks = df_disp_matriks[df_disp_matriks["NAMA DAPUR"] == sel_dapur_filter]
        if sel_sup_filter != "Semua Supplier":
            df_disp_matriks = df_disp_matriks[df_disp_matriks["NAMA SUPPLIER"] == sel_sup_filter]
            
        st.dataframe(df_disp_matriks, use_container_width=True, hide_index=True)

# ---------------------------------------------------------
# MODUL 6: SCORING & EVALUASI
# ---------------------------------------------------------
elif menu == "🎯 Scoring & Evaluasi":
    st.subheader("🎯 Supplier Scoring & Evaluation Engine")
    
    st.markdown("##### ⚙️ Pengaturan Bobot Evaluasi")
    w_col1, w_col2, w_col3, w_col4 = st.columns(4)
    with w_col1:
        w_harga = st.slider("Bobot Harga (%)", 0, 100, 40)
    with w_col2:
        w_jarak = st.slider("Bobot Jarak (%)", 0, 100, 20)
    with w_col3:
        w_rating = st.slider("Bobot Rating (%)", 0, 100, 20)
    with w_col4:
        w_ketepatan = st.slider("Bobot Ketepatan (%)", 0, 100, 20)
        
    total_w = w_harga + w_jarak + w_rating + w_ketepatan
    if total_w != 100:
        st.warning(f"⚠️ Total bobot saat ini: {total_w}%. Disarankan total bobot bernilai 100%.")

    df_sup = st.session_state["df_supplier_state"].copy()
    df_harga = st.session_state["df_harga_bb"].copy()
    
    if df_sup.empty:
        st.warning("Data supplier masih kosong.")
    else:
        scores = []
        for _, sup in df_sup.iterrows():
            k_sup = sup.get("KODE SUPPLIER", "")
            
            sup_harga = df_harga[df_harga["KODE SUPPLIER"] == k_sup]
            avg_harga = sup_harga["HARGA PENAWARAN"].mean() if not sup_harga.empty else 0
            
            score_harga = max(0, 100 - (avg_harga / 1000)) if avg_harga > 0 else 50
            
            score_rating = float(sup.get("RATING", 4.0)) * 20
            score_ketepatan = float(sup.get("KETEPATAN_KIRIM", 85))
            score_jarak = 80
            
            final_score = (
                (score_harga * w_harga / 100) +
                (score_jarak * w_jarak / 100) +
                (score_rating * w_rating / 100) +
                (score_ketepatan * w_ketepatan / 100)
            )
            
            scores.append({
                "KODE SUPPLIER": k_sup,
                "NAMA SUPPLIER": sup.get("NAMA SUPPLIER", "-"),
                "SCORE HARGA": round(score_harga, 1),
                "SCORE RATING": round(score_rating, 1),
                "SCORE KETEPATAN": round(score_ketepatan, 1),
                "FINAL SCORE": round(final_score, 2)
            })
            
        df_scoring = pd.DataFrame(scores).sort_values(by="FINAL SCORE", ascending=False)
        st.markdown("##### 🏆 Ranking Hasil Evaluasi Supplier")
        st.dataframe(df_scoring, use_container_width=True, hide_index=True)

# ---------------------------------------------------------
# MODUL 7: HET & KOMPARASI PASAR
# ---------------------------------------------------------
elif menu == "📈 HET & Komparasi Pasar":
    st.subheader("📈 Monitoring HET & Komparasi Harga Pasar")
    
    df_bb = st.session_state["df_bb_state"].copy()
    df_harga = st.session_state["df_harga_bb"].copy()
    
    if df_harga.empty:
        st.info("Belum ada penawaran harga masuk dari supplier.")
    else:
        df_comp = pd.merge(df_harga, df_bb[["ITEM BB", "HET"]], on="ITEM BB", how="left")
        df_comp["HET"] = df_comp["HET"].fillna(0)
        df_comp["DEViasi HET (%)"] = np.where(
            df_comp["HET"] > 0,
            ((df_comp["HARGA PENAWARAN"] - df_comp["HET"]) / df_comp["HET"]) * 100,
            0
        )
        
        st.markdown("##### 🛒 Komparasi Harga Penawaran vs HET Master")
        
        def highlight_het(val):
            color = 'red' if val > 0 else 'green'
            return f'color: {color}'

        st.dataframe(
            df_comp[["KODE SUPPLIER", "NAMA SUPPLIER", "ITEM BB", "SATUAN", "HARGA PENAWARAN", "HET", "DEViasi HET (%)"]],
            use_container_width=True,
            hide_index=True
        )
