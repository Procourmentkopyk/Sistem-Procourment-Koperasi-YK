import streamlit as st
import pandas as pd
import numpy as np
import os

# Konfigurasi Halaman
st.set_page_config(
    page_title="SPPG Procurement Engine | Koperasi YK",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

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
    
    /* Custom Styling Baris Tabel Selang-Seling */
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

# Inisialisasi Data Dapur di Session State
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

# --- SIDEBAR NAVIGASI ---
with st.sidebar:
    st.markdown("### ⚡ **SPPG Engine**")
    st.caption("Koperasi YK")
    st.markdown("---")
    
    st.subheader("🚀 Modul Operasional")
    
    modul_options = [
        "📊 Dashboard & HET",
        "🏬 Kelola Data Dapur",
        "💬 WA & PO Generator",
        "🚛 Matriks Jarak",
        "🎯 Scoring & Evaluasi",
        "📈 HET & Komparasi Pasar"
    ]
    
    menu = st.radio("Pilih Modul:", modul_options, index=1)
    st.markdown("---")
    st.caption(f"Status Data: 🟢 {file_status}")

# --- HEADER UTAMA ---
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
# MODUL 2: KELOLA DATA DAPUR SPPG
# ---------------------------------------------------------
if menu == "🏬 Kelola Data Dapur":
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
        # Header Tabel
        h_cols = st.columns([0.8, 1.8, 2.5, 1.0, 1.1, 1.1, 0.9, 0.5, 0.5])
        headers = ["KODE", "NAMA DAPUR", "ALAMAT", "PIC", "LATITUDE", "LONGITUDE", "MAPS", "EDIT", "HAPUS"]
        
        for col, h in zip(h_cols, headers):
            col.markdown(f"**{h}**")
        st.markdown("<hr style='margin-top:0; margin-bottom:10px;'>", unsafe_allow_html=True)
        
        # Iterasi Baris Data dengan Warna Selang-Seling
        for i, r in df_dapur.iterrows():
            # Tentukan warna latar belakang berdasarkan baris genap/ganjil
            bg_color = "#f8fafc" if i % 2 == 0 else "#ffffff"
            
            # Buat kontainer bergaya kartu tipis untuk setiap baris
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
                
                # Direct Link Google Maps
                if pd.notna(lat_val) and pd.notna(lon_val):
                    gmap_url = f"https://www.google.com/maps/search/?api=1&query={lat_val},{lon_val}"
                    c_map.markdown(f"[📍 Buka]({gmap_url})")
                else:
                    c_map.write("-")
                
                # Tombol Aksi
                if c_edit.button("✏️", key=f"edit_{i}"):
                    open_edit_dapur_dialog(i)
                    
                if c_del.button("🗑️", key=f"del_{i}"):
                    open_delete_dapur_dialog(i)
                
                st.markdown("</div>", unsafe_allow_html=True)
    else:
        st.info("Belum ada data dapur. Klik 'Tambah Dapur Baru' untuk menambahkan.")

# ---------------------------------------------------------
# MODUL LAINNYA
# ---------------------------------------------------------
elif menu == "📊 Dashboard & HET":
    st.subheader("📊 Dashboard Utama & Pemantauan HET")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Dapur SPPG", len(st.session_state["df_dapur_state"]), delta="Active")
    m2.metric("Total Supplier", "45 Supplier", delta="+3 Bulan Ini")
    m3.metric("Kategori Barang", "6 Kategori", delta="Lengkap")
    m4.metric("Rata-rata Ketepatan", "94.2%", delta="+1.5%")

else:
    st.subheader(f"Modul {menu}")
    st.info("Modul sedang aktif dan dapat diakses dari sidebar sebelah kiri.")
