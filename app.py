import streamlit as st
import pandas as pd
import numpy as np
import os

# Konfigurasi Halaman Clean Mode
st.set_page_config(
    page_title="SPPG Procurement Engine | Koperasi YK",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS
st.markdown("""
<style>
    .header-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .status-badge {
        display: inline-block;
        padding: 0.35em 0.65em;
        font-size: 80%;
        font-weight: 600;
        border-radius: 20px;
        color: #065f46;
        background-color: #d1fae5;
    }
    div.stButton > button {
        border-radius: 8px;
        transition: all 0.2s ease-in-out;
    }
    div.stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
    }
</style>
""", unsafe_allow_html=True)

# Function Load Data Excel / XLSB
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

# Inisialisasi Session State Modul
if "current_modul" not in st.session_state:
    st.session_state["current_modul"] = "2. Kelola Data Dapur SPPG"

# Inisialisasi & Cleaning Data Dapur (Latitude/Longitude Tetap Ada)
if "df_dapur_state" not in st.session_state:
    if data_excel and isinstance(data_excel, dict) and "Data Dapur" in data_excel:
        raw_df = data_excel["Data Dapur"].copy()
        
        # Deteksi Header jika terbungkus Unnamed
        if "Unnamed" in str(raw_df.columns[0]):
            header_idx = raw_df[raw_df.apply(lambda row: row.astype(str).str.contains('NAMA DAPUR').any(), axis=1)].index
            if not header_idx.empty:
                idx = header_idx[0]
                raw_df.columns = raw_df.iloc[idx]
                raw_df = raw_df.iloc[idx+1:].reset_index(drop=True)
                raw_df = raw_df.dropna(how="all", subset=["NAMA DAPUR"])
        
        # Bersihkan nama kolom
        raw_df.columns = [str(c).strip().upper() for c in raw_df.columns]
        
        # Kolom yang dipertahankan
        cols_to_keep = [c for c in ["KODE", "NAMA DAPUR", "ALAMAT", "PIC", "LATITUDE", "LONGTITUDE", "LONGITUDE", "KOTA/KABUPATEN"] if c in raw_df.columns]
        df_clean = raw_df[cols_to_keep].copy()
        
        # Standardisasi Nama Kolom Longitude
        if "LONGTITUDE" in df_clean.columns:
            df_clean = df_clean.rename(columns={"LONGTITUDE": "LONGITUDE"})
            
        # Konversi Lat/Lon ke Angka/Float
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

# --- POP-UP DIALOG: TAMBAH DAPUR ---
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

# --- POP-UP DIALOG: EDIT DAPUR ---
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

# --- POP-UP DIALOG: KONFIRMASI HAPUS DAPUR ---
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

# --- HEADER UTAMA ---
st.markdown(f"""
<div class="header-card">
    <div style="display: flex; justify-content: space-between; align-items: center;">
        <div>
            <h2 style="margin:0; color: white;">🏭 Enterprise Procurement & Evaluation System</h2>
            <p style="margin:0; opacity: 0.8;">Koperasi YK — Sistem Evaluasi Supplier & Dapur SPPG</p>
        </div>
        <div>
            <span class="status-badge">🟢 Data: {file_status}</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# --- NAVIGASI MODUL ---
st.markdown("##### 🚀 **Pilih Modul Operasional:**")
nav_col1, nav_col2, nav_col3, nav_col4, nav_col5, nav_col6 = st.columns(6)

with nav_col1:
    if st.button("📊 Dashboard", use_container_width=True):
        st.session_state["current_modul"] = "1. Dashboard & HET"
with nav_col2:
    if st.button("🏬 Kelola Dapur", use_container_width=True):
        st.session_state["current_modul"] = "2. Kelola Data Dapur SPPG"
with nav_col3:
    if st.button("💬 WA & PO", use_container_width=True):
        st.session_state["current_modul"] = "3. Penawaran WA & Update Harga PO"
with nav_col4:
    if st.button("🚛 Matriks Jarak", use_container_width=True):
        st.session_state["current_modul"] = "4. Data Supplier & Matriks Jarak"
with nav_col5:
    if st.button("🎯 Scoring", use_container_width=True):
        st.session_state["current_modul"] = "5. Penilaian & Penentuan Supplier per Dapur"
with nav_col6:
    if st.button("📈 HET & Pasar", use_container_width=True):
        st.session_state["current_modul"] = "6. Analisis Komparasi Harga Pasar & HET"

st.markdown("---")

menu = st.session_state["current_modul"]

# ---------------------------------------------------------
# MODUL 2: KELOLA DATA DAPUR SPPG
# ---------------------------------------------------------
if menu == "2. Kelola Data Dapur SPPG":
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
        # Standardisasi kolom lat/lon khusus untuk st.map
        map_df = pd.DataFrame({
            "lat": map_data["LATITUDE"].astype(float),
            "lon": map_data["LONGITUDE"].astype(float)
        })
        st.map(map_df, zoom=10, use_container_width=True)
    else:
        st.warning("Data Koordinat (Latitude/Longitude) belum terisi dengan benar.")
        
    st.markdown("---")
    
    # --- TABEL DATA DAPUR DENGAN KOORDINAT & LINK GOOGLE MAPS ---
    if not df_dapur.empty:
        # Header Kolom
        cols = st.columns([0.8, 1.8, 2.5, 1.0, 1.2, 1.2, 1.0, 0.5, 0.5])
        headers = ["KODE", "NAMA DAPUR", "ALAMAT", "PIC", "LATITUDE", "LONGITUDE", "MAPS", "EDIT", "HAPUS"]
        
        for col, h in zip(cols, headers):
            col.markdown(f"**{h}**")
        st.markdown("---")
        
        # Baris Data
        for i, r in df_dapur.iterrows():
            c_kode, c_nama, c_alamat, c_pic, c_lat, c_lon, c_map, c_edit, c_del = st.columns([0.8, 1.8, 2.5, 1.0, 1.2, 1.2, 1.0, 0.5, 0.5])
            
            lat_val = r.get("LATITUDE", None)
            lon_val = r.get("LONGITUDE", None)
            
            c_kode.write(r.get("KODE", "-"))
            c_nama.write(r.get("NAMA DAPUR", "-"))
            c_alamat.write(r.get("ALAMAT", "-"))
            c_pic.write(r.get("PIC", "-"))
            c_lat.write(f"{lat_val:.5f}" if pd.notna(lat_val) else "-")
            c_lon.write(f"{lon_val:.5f}" if pd.notna(lon_val) else "-")
            
            # Direct Link Google Maps
            if pd.notna(lat_val) and pd.notna(lon_val):
                gmap_url = f"https://www.google.com/maps/search/?api=1&query={lat_val},{lon_val}"
                c_map.markdown(f"[📍 Buka]({gmap_url})")
            else:
                c_map.write("-")
            
            # Tombol Edit
            if c_edit.button("✏️", key=f"edit_{i}"):
                open_edit_dapur_dialog(i)
                
            # Tombol Hapus
            if c_del.button("🗑️", key=f"del_{i}"):
                open_delete_dapur_dialog(i)
    else:
        st.info("Belum ada data dapur. Klik 'Tambah Dapur Baru' untuk menambahkan.")

# ---------------------------------------------------------
# MODUL LAINNYA
# ---------------------------------------------------------
elif menu == "1. Dashboard & HET":
    st.subheader("📊 Dashboard Utama & Pemantauan HET")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Dapur SPPG", len(st.session_state["df_dapur_state"]), delta="Active")
    m2.metric("Total Supplier", "45 Supplier", delta="+3 Bulan Ini")
    m3.metric("Kategori Barang", "6 Kategori", delta="Lengkap")
    m4.metric("Rata-rata Ketepatan", "94.2%", delta="+1.5%")
