import streamlit as st
import pandas as pd
import numpy as np
import os
import io
import uuid
import urllib.parse
from datetime import datetime

# --- KONFIGURASI HALAMAN ---
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

    .wa-card {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 20px;
        margin-top: 10px;
    }
    .wa-button {
        background-color: #25D366;
        color: white !important;
        padding: 14px 20px;
        border-radius: 10px;
        text-align: center;
        font-weight: bold;
        font-size: 16px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        display: block;
        text-decoration: none;
        margin-top: 15px;
    }
    .wa-button:hover {
        background-color: #1eb854;
        text-decoration: none;
    }
</style>
""", unsafe_allow_html=True)

# URL Web App Deployment dari Google Apps Script
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwzJCP1oFC1lFgLqc6NogtIm4ClzspiNUj6M54gp7dOV1_9YImsuH4nskao2vLXe9OkJA/exec"

# --- HELPER FUNCTIONS ---
def normalisasi(text):
    return str(text or '').strip().lower()

def buat_token():
    return uuid.uuid4().hex

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

# =========================================================
# INISIALISASI SESSION STATE DATA
# =========================================================

# 1. Inisialisasi Data Dapur di Session State
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

# 2. Inisialisasi Data Supplier Link di Session State
if "df_supplier_state" not in st.session_state:
    if data_excel and isinstance(data_excel, dict) and "Supplier Link" in data_excel:
        df_sup_link = data_excel["Supplier Link"].dropna(how="all").copy()
        if "TOKEN" not in df_sup_link.columns:
            df_sup_link["TOKEN"] = [buat_token() for _ in range(len(df_sup_link))]
        if "NO WA" not in df_sup_link.columns:
            df_sup_link["NO WA"] = "6285169796974"
        st.session_state["df_supplier_state"] = df_sup_link.reset_index(drop=True)
    else:
        st.session_state["df_supplier_state"] = pd.DataFrame({
            "KODE SUPPLIER": ["SUP-001", "SUP-002"],
            "NAMA SUPPLIER": ["UD Berkah Tani", "CV Sayur Segar"],
            "TOKEN": [buat_token(), buat_token()],
            "NO WA": ["6285169796974", "6281234567890"],
            "JENIS BB 1": ["Sayuran", "Daging"],
            "JENIS BB 2": ["Bumbu", "Ayam"]
        })

# 3. Inisialisasi Update Harga BB (Master Bahan Baku)
if "df_harga_bb" not in st.session_state:
    sheet_name = "UPDATE HARGA BB" if (data_excel and "UPDATE HARGA BB" in data_excel) else "Harga Supplier"
    if data_excel and isinstance(data_excel, dict) and sheet_name in data_excel:
        st.session_state["df_harga_bb"] = data_excel[sheet_name].copy()
    else:
        st.session_state["df_harga_bb"] = pd.DataFrame({
            "PN": [1, 2, 3, 4, 5],
            "KATEGORI": ["Sembako", "Daging", "Sembako", "Minyak", "Bumbu"],
            "JENIS BB": ["Beras", "Daging", "Telur", "Minyak", "Bawang"],
            "ITEM BB": ["Beras Medium", "Daging Ayam Ras", "Telur Ayam Ras", "Minyak Goreng", "Bawang Merah"],
            "SATUAN DAPUR": ["Kg", "Kg", "Kg", "Liter", "Kg"],
            "TARGET HARGA": [12000, 35000, 27000, 16000, 30000]
        })

# 4. Inisialisasi Template Pesan WA
if "wa_template_config" not in st.session_state:
    st.session_state["wa_template_config"] = {
        "header_admin": "Procurement Team Koperasi YK",
        "wa_admin": "085169796974",
        "periode_default": f"Minggu ke-{datetime.now().isocalendar()[1]} ({datetime.now().strftime('%B %Y')})",
        "template_text": """Yth. *{NAMA_SUPPLIER}*,

Kami dari *Koperasi YK (SPPG Procurement)* memohon informasi penawaran update harga bahan baku terbaru untuk periode *{PERIODE_MINGGU}*.

Berikut daftar item bahan baku yang dapat Anda tawarkan minggu ini:
{DAFTAR_BAHAN}
Mohon untuk mengisi update harga penawaran Anda melalui Form Online resmi berikut:
🔗 {LINK_FORM}

*Catatan:*
- Pengisian paling lambat dilakukan hari ini.
- Harga yang dimasukkan adalah harga satuan sesuai spesifikasi.

Terima kasih atas kerja samanya.

Salam,
*{HEADER_ADMIN}*
WA Admin: {WA_ADMIN}"""
    }

# =========================================================
# POP-UP DIALOG DAPUR
# =========================================================

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

# --- POP-UP DIALOG: SUPPLIER ---
@st.dialog("➕ Tambah Supplier Baru")
def open_add_supplier_dialog():
    with st.form("form_add_sup"):
        c1, c2 = st.columns(2)
        with c1:
            kode = st.text_input("Kode Supplier (Contoh: SUP-003)")
            nama = st.text_input("Nama Supplier")
            no_wa = st.text_input("Nomor WA (Format: 628xxx)", value="6285169796974")
        with c2:
            j1 = st.text_input("Jenis BB 1 (Contoh: Sayuran)")
            j2 = st.text_input("Jenis BB 2 (Opsional)")
            j3 = st.text_input("Jenis BB 3 (Opsional)")
            
        if st.form_submit_button("✨ Simpan Supplier", use_container_width=True):
            token = buat_token()
            new_row = pd.DataFrame([{
                "KODE SUPPLIER": kode,
                "NAMA SUPPLIER": nama,
                "TOKEN": token,
                "LINK FORM": f"{WEB_APP_URL}?token={token}",
                "NO WA": no_wa,
                "JENIS BB 1": j1,
                "JENIS BB 2": j2,
                "JENIS BB 3": j3
            }])
            st.session_state["df_supplier_state"] = pd.concat([st.session_state["df_supplier_state"], new_row], ignore_index=True)
            st.success("Supplier baru berhasil ditambahkan!")
            st.rerun()

@st.dialog("✏️ Edit Data Supplier")
def open_edit_supplier_dialog(index):
    df = st.session_state["df_supplier_state"]
    row = df.iloc[index]
    
    with st.form("form_edit_sup"):
        c1, c2 = st.columns(2)
        with c1:
            kode = st.text_input("Kode Supplier", value=str(row.get("KODE SUPPLIER", "")))
            nama = st.text_input("Nama Supplier", value=str(row.get("NAMA SUPPLIER", "")))
            no_wa = st.text_input("Nomor WA", value=str(row.get("NO WA", "6285169796974")))
        with c2:
            j1 = st.text_input("Jenis BB 1", value=str(row.get("JENIS BB 1", "")))
            j2 = st.text_input("Jenis BB 2", value=str(row.get("JENIS BB 2", "")))
            j3 = st.text_input("Jenis BB 3", value=str(row.get("JENIS BB 3", "")))
            
        if st.form_submit_button("💾 Update Supplier", use_container_width=True):
            st.session_state["df_supplier_state"].at[index, "KODE SUPPLIER"] = kode
            st.session_state["df_supplier_state"].at[index, "NAMA SUPPLIER"] = nama
            st.session_state["df_supplier_state"].at[index, "NO WA"] = no_wa
            st.session_state["df_supplier_state"].at[index, "JENIS BB 1"] = j1
            st.session_state["df_supplier_state"].at[index, "JENIS BB 2"] = j2
            st.session_state["df_supplier_state"].at[index, "JENIS BB 3"] = j3
            st.success("Data supplier berhasil diperbarui!")
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
        "🤝 Data Supplier & Link Form",
        "💬 WA & PO Generator",
        "🚛 Matriks Jarak",
        "🎯 Scoring & Evaluasi",
        "📈 HET & Komparasi Pasar"
    ]
    
    menu = st.radio("Pilih Modul:", modul_options, index=1)
    st.markdown("---")
    st.caption(f"Status Data: 🟢 {file_status}")

# --- HEADER UTAMA ---
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
    m1.metric("Total Dapur SPPG", len(st.session_state["df_dapur_state"]), delta="Active")
    m2.metric("Total Supplier", f"{len(st.session_state['df_supplier_state'])} Supplier", delta="Aktif")
    m3.metric("Item Bahan Baku Master", len(st.session_state["df_harga_bb"]))
    m4.metric("Status Integrasi WA", "🟢 Siap Digunakan")

    st.markdown("---")
    st.markdown("##### 📈 **Ringkasan Master Bahan Baku & HET Target**")
    st.dataframe(st.session_state["df_harga_bb"], use_container_width=True)

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
        # Header Tabel
        h_cols = st.columns([0.8, 1.8, 2.5, 1.0, 1.1, 1.1, 0.9, 0.5, 0.5])
        headers = ["KODE", "NAMA DAPUR", "ALAMAT", "PIC", "LATITUDE", "LONGITUDE", "MAPS", "EDIT", "HAPUS"]
        
        for col, h in zip(h_cols, headers):
            col.markdown(f"**{h}**")
        st.markdown("<hr style='margin-top:0; margin-bottom:10px;'>", unsafe_allow_html=True)
        
        # Iterasi Baris Data dengan Warna Selang-Seling
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
# MODUL 3: DATA SUPPLIER & LINK FORM
# ---------------------------------------------------------
elif menu == "🤝 Data Supplier & Link Form":
    st.subheader("🤝 Kelola Data Supplier & Link Web App")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        if st.button("➕ Tambah Supplier Baru", type="primary", use_container_width=True):
            open_add_supplier_dialog()
            
    with col2:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
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
    df_sup = st.session_state["df_supplier_state"].copy()
    
    if not df_sup.empty:
        for idx, row in df_sup.iterrows():
            tkn = row.get("TOKEN", "")
            form_url = f"{WEB_APP_URL}?token={tkn}" if pd.notna(tkn) else ""
            
            with st.expander(f"🤝 **{row.get('KODE SUPPLIER', '')} - {row.get('NAMA SUPPLIER', '')}** (WA: +{row.get('NO WA', '-')})"):
                c1, c2, c3 = st.columns([3, 1, 1])
                with c1:
                    st.write(f"**Jenis BB:** {row.get('JENIS BB 1', '-')} / {row.get('JENIS BB 2', '-')} / {row.get('JENIS BB 3', '-')}")
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
        st.info("Belum ada data supplier.")

# ---------------------------------------------------------
# MODUL 4: WA & PO GENERATOR (REQUEST HARGA 1-KLIK)
# ---------------------------------------------------------
elif menu == "💬 WA & PO Generator":
    st.subheader("💬 Request Update Harga via WhatsApp (1-Klik)")
    
    df_sup = st.session_state["df_supplier_state"]
    df_bb = st.session_state["df_harga_bb"]
    cfg = st.session_state["wa_template_config"]
    
    # Pengaturan Template Mingguan
    with st.expander("⚙️ **Pengaturan Master Template Pesan WA Mingguan (Klik untuk Mengubah)**", expanded=False):
        c_tp1, c_tp2 = st.columns(2)
        with c_tp1:
            st.session_state["wa_template_config"]["periode_default"] = st.text_input(
                "Periode Minggu Pengiriman:", value=cfg["periode_default"]
            )
            st.session_state["wa_template_config"]["header_admin"] = st.text_input(
                "Nama Pengirim / Divisi:", value=cfg["header_admin"]
            )
        with c_tp2:
            st.session_state["wa_template_config"]["wa_admin"] = st.text_input(
                "Nomor WA Admin:", value=cfg["wa_admin"]
            )
            
        st.session_state["wa_template_config"]["template_text"] = st.text_area(
            "Format Master Pesan WA:", value=cfg["template_text"], height=180
        )

    st.markdown("---")
    
    col_left, col_right = st.columns([1, 1])
    
    with col_left:
        st.markdown("##### 📱 **1. Pilih Supplier Tujuan**")
        
        if not df_sup.empty:
            sup_names = df_sup["NAMA SUPPLIER"].tolist()
            selected_sup_name = st.selectbox("Pilih Supplier Target:", sup_names)
            
            sup_row = df_sup[df_sup["NAMA SUPPLIER"] == selected_sup_name].iloc[0]
            kode_sup = str(sup_row.get("KODE SUPPLIER", ""))
            token_sup = str(sup_row.get("TOKEN", ""))
            
            target_wa = str(sup_row.get("NO WA", "6285169796974")).replace("-", "").replace(" ", "").replace("+", "")
            if target_wa.startswith("0"):
                target_wa = "62" + target_wa[1:]
                
            form_link = f"{WEB_APP_URL}?token={token_sup}"
            
            jenis_bb_list = []
            for col in sup_row.index:
                if "JENIS" in str(col).upper() and pd.notna(sup_row[col]) and str(sup_row[col]).strip() != "":
                    jenis_bb_list.append(normalisasi(sup_row[col]))
        else:
            selected_sup_name = "Supplier Example"
            kode_sup = "SUP-001"
            target_wa = "6285169796974"
            form_link = WEB_APP_URL
            jenis_bb_list = []

        draf_list_bahan = ""
        item_count = 0
        
        col_item = "ITEM BB" if "ITEM BB" in df_bb.columns else "NAMA BAHAN BAKU"
        col_satuan = "SATUAN DAPUR" if "SATUAN DAPUR" in df_bb.columns else "SATUAN"
        col_jenis = "JENIS BB" if "JENIS BB" in df_bb.columns else "KATEGORI"
        
        for idx, r in df_bb.iterrows():
            item_j = normalisasi(r.get(col_jenis, ""))
            if not jenis_bb_list or any(j in item_j for j in jenis_bb_list):
                item_count += 1
                b_name = r.get(col_item, f"Item {item_count}")
                b_sat = r.get(col_satuan, "Kg")
                draf_list_bahan += f"{item_count}. {b_name} ({b_sat})\n"

        if item_count == 0:
            draf_list_bahan = "(Semua Jenis Bahan Baku Operasional)\n"

        pesan_jadi = cfg["template_text"].format(
            NAMA_SUPPLIER=selected_sup_name,
            PERIODE_MINGGU=cfg["periode_default"],
            LINK_FORM=form_link,
            DAFTAR_BAHAN=draf_list_bahan,
            HEADER_ADMIN=cfg["header_admin"],
            WA_ADMIN=cfg["wa_admin"]
        )

        st.info(f"📌 **Target**: {selected_sup_name} ({kode_sup})\n\n📞 **WA**: +{target_wa}\n\n📋 **Total Item**: {item_count} Bahan Baku")

    with col_right:
        st.markdown("##### ✏️ **2. Pratinjau Teks & Kirim**")
        final_message = st.text_area("Pratinjau Pesan WA (Bisa Diedit):", value=pesan_jadi, height=320)
        
        encoded_message = urllib.parse.quote(final_message)
        wa_link = f"https://wa.me/{target_wa}?text={encoded_message}"
        
        st.markdown(f"""
        <div class="wa-card">
            <p style="margin:0; font-weight:bold; color:#1e293b;">🚀 Langsung Kirim Ke WhatsApp Supplier</p>
            <p style="margin:5px 0 0 0; font-size:13px; color:#64748b;">Aplikasi akan membuka WhatsApp secara otomatis dengan nomor tujuan dan pesan terformat di atas.</p>
            <a href="{wa_link}" target="_blank" class="wa-button">
                📲 KIRIM PESAN WA SEKARANG (1-KLIK)
            </a>
        </div>
        """, unsafe_allow_html=True)

# ---------------------------------------------------------
# MODUL LAINNYA
# ---------------------------------------------------------
else:
    st.subheader(f"Modul {menu}")
    st.info("Modul sedang aktif dan dapat diakses dari sidebar sebelah kiri.")
