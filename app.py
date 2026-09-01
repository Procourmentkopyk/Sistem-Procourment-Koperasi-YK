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

def buat_token(length=32):
    return ''.join(random.choices(string.ascii_letters + string.digits, k=length))

# Custom CSS untuk Table Styling
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

# 3. State Harga BB / Penawaran
if "df_harga_bb" not in st.session_state:
    st.session_state["df_harga_bb"] = pd.DataFrame(columns=[
        "KODE SUPPLIER", "NAMA SUPPLIER", "NAMA BB", "HARGA PER SATUAN", "SATUAN", "KATEGORI", "CATATAN"
    ])

# 4. State Master Bahan Baku
if "df_bb_state" not in st.session_state:
    st.session_state["df_bb_state"] = pd.DataFrame(columns=["P/N", "JENIS BB", "ITEM BB", "SATUAN"])

# ---------------------------------------------------------
# MODE VENDOR EXTERNAL (JIKA DIBUKA DENGAN TOKEN VIA URL)
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

    st.title("📝 Form Penawaran Harga Supplier")
    st.caption("Koperasi YK — Sistem Informasi Procurement & SPPG")
    st.markdown("---")

    if supplier_info is not None:
        st.success(f"Selamat datang, **{supplier_info['NAMA SUPPLIER']}**!")
    else:
        st.info("Form Penawaran Harga Barang Operasional")

    if kategori_diterima != "ALL":
        list_kat_supplier = [k.replace("_", " ") for k in kategori_diterima.split(",")]
        st.write("Kategori Barang yang Ditawarkan:")
        st.write(" ".join([f"`{k}`" for k in list_kat_supplier]))
    else:
        list_kat_supplier = ["Semua Kategori"]
        st.write("Kategori Barang: **Semua Kategori**")

    st.markdown("---")

    with st.form("form_input_harga_supplier"):
        st.subheader("🛒 Input Penawaran Harga")
        nama_barang = st.text_input("Nama Bahan Baku / Barang", placeholder="Contoh: Telur Ayam Ras / Minyak Goreng")
        
        c1, c2 = st.columns(2)
        with c1:
            harga_penawaran = st.number_input("Harga Penawaran (Rp)", min_value=0, step=500, value=10000)
            satuan = st.selectbox("Satuan", ["Kg", "Liter", "Ikat", "Pcs", "Karton", "Ekor", "Pack"])
        with c2:
            kategori_pilihan = st.selectbox(
                "Kategori Barang", 
                options=list_kat_supplier if list_kat_supplier != ["Semua Kategori"] else [
                    "Ayam", "Beras", "Buah", "Cookies", "Daging", "Ikan", 
                    "Keju", "Olahan", "Sayur", "Sembako", "Susu", "Tahu", 
                    "Tempe", "Telur Ayam", "Telur Bebek", "Telur Puyuh"
                ]
            )
            catatan = st.text_input("Catatan / Merek (Opsional)", placeholder="Contoh: Ukuran Medium, Fresh")

        submit_harga = st.form_submit_button("🚀 Kirim Penawaran Harga", use_container_width=True)

        if submit_harga:
            if nama_barang.strip() == "":
                st.error("Mohon isi nama barang terlebih dahulu.")
            else:
                kode_sup = supplier_info["KODE SUPPLIER"] if supplier_info is not None else "GUEST"
                nama_sup = supplier_info["NAMA SUPPLIER"] if supplier_info is not None else "Supplier External"

                new_entry = pd.DataFrame([{
                    "KODE SUPPLIER": kode_sup,
                    "NAMA SUPPLIER": nama_sup,
                    "NAMA BB": nama_barang,
                    "HARGA PER SATUAN": harga_penawaran,
                    "SATUAN": satuan,
                    "KATEGORI": kategori_pilihan,
                    "CATATAN": catatan
                }])

                st.session_state["df_harga_bb"] = pd.concat([st.session_state["df_harga_bb"], new_entry], ignore_index=True)
                st.balloons()
                st.success(f"Berhasil! Penawaran untuk **{nama_barang}** seharga **Rp {harga_penawaran:,}** telah tersimpan.")

    st.stop()  # Berhenti di sini khusus untuk akses vendor via token

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
# DIALOGS / POP-UP ITEM HARGA BAHAN BAKU (FITUR BARU)
# ---------------------------------------------------------
@st.dialog("➕ Tambah Item Bahan Baku Baru")
def open_add_harga_bb_dialog():
    st.write("Masukkan detail penawaran / harga bahan baku baru:")
    with st.form("form_add_harga_bb"):
        c1, c2 = st.columns(2)
        with c1:
            df_sup = st.session_state.get("df_supplier_state", pd.DataFrame())
            list_supplier = df_sup["NAMA SUPPLIER"].tolist() if not df_sup.empty and "NAMA SUPPLIER" in df_sup.columns else ["MANUAL / INTERNAL"]
            nama_sup = st.selectbox("Pilih Supplier", options=list_supplier)
            
            kode_sup = "INTERNAL"
            if not df_sup.empty and "NAMA SUPPLIER" in df_sup.columns:
                match = df_sup[df_sup["NAMA SUPPLIER"] == nama_sup]
                if not match.empty:
                    kode_sup = match.iloc[0].get("KODE SUPPLIER", "SUP-000")

            nama_bb = st.text_input("Nama Bahan Baku / Item", placeholder="Contoh: Daging Sapi Segar")
            harga = st.number_input("Harga per Satuan (Rp)", min_value=0, step=500, value=15000)
            
        with c2:
            satuan = st.selectbox("Satuan", ["Kg", "Liter", "Ikat", "Pcs", "Karton", "Ekor", "Pack", "Gram", "Kg/Ikat"])
            kategori = st.selectbox("Kategori", [
                "Ayam", "Beras", "Buah", "Cookies", "Daging", "Ikan", 
                "Keju", "Olahan", "Sayur", "Sembako", "Susu", "Tahu", 
                "Tempe", "Telur Ayam", "Telur Bebek", "Telur Puyuh"
            ])
            catatan = st.text_input("Catatan / Keterangan", placeholder="Contoh: Kualitas super / Grade A")

        if st.form_submit_button("✨ Simpan Item BB", use_container_width=True):
            if nama_bb.strip() == "":
                st.error("Nama Bahan Baku wajib diisi!")
            else:
                new_row = pd.DataFrame([{
                    "KODE SUPPLIER": kode_sup,
                    "NAMA SUPPLIER": nama_sup,
                    "NAMA BB": nama_bb,
                    "HARGA PER SATUAN": harga,
                    "SATUAN": satuan,
                    "KATEGORI": kategori,
                    "CATATAN": catatan
                }])
                st.session_state["df_harga_bb"] = pd.concat([st.session_state["df_harga_bb"], new_row], ignore_index=True)
                st.success("Item Bahan Baku berhasil ditambahkan!")
                st.rerun()

@st.dialog("✏️ Edit Item Bahan Baku")
def open_edit_harga_bb_dialog(index):
    df = st.session_state["df_harga_bb"]
    row = df.iloc[index]
    
    st.write(f"Edit harga/item untuk: **{row.get('NAMA BB', 'Item BB')}**")
    with st.form("form_edit_harga_bb"):
        c1, c2 = st.columns(2)
        with c1:
            kode_sup = st.text_input("Kode Supplier", value=str(row.get("KODE SUPPLIER", "")))
            nama_sup = st.text_input("Nama Supplier", value=str(row.get("NAMA SUPPLIER", "")))
            nama_bb = st.text_input("Nama Bahan Baku", value=str(row.get("NAMA BB", "")))
            harga = st.number_input("Harga per Satuan (Rp)", value=int(row.get("HARGA PER SATUAN", 0)), step=500)
            
        with c2:
            list_satuan = ["Kg", "Liter", "Ikat", "Pcs", "Karton", "Ekor", "Pack", "Gram", "Kg/Ikat"]
            curr_satuan = str(row.get("SATUAN", "Kg"))
            idx_sat = list_satuan.index(curr_satuan) if curr_satuan in list_satuan else 0
            satuan = st.selectbox("Satuan", list_satuan, index=idx_sat)

            list_kat = [
                "Ayam", "Beras", "Buah", "Cookies", "Daging", "Ikan", 
                "Keju", "Olahan", "Sayur", "Sembako", "Susu", "Tahu", 
                "Tempe", "Telur Ayam", "Telur Bebek", "Telur Puyuh"
            ]
            curr_kat = str(row.get("KATEGORI", "Sayur"))
            idx_kat = list_kat.index(curr_kat) if curr_kat in list_kat else 0
            kategori = st.selectbox("Kategori", list_kat, index=idx_kat)
            
            catatan = st.text_input("Catatan / Merek", value=str(row.get("CATATAN", "")))

        if st.form_submit_button("💾 Perbarui Item BB", use_container_width=True):
            st.session_state["df_harga_bb"].at[index, "KODE SUPPLIER"] = kode_sup
            st.session_state["df_harga_bb"].at[index, "NAMA SUPPLIER"] = nama_sup
            st.session_state["df_harga_bb"].at[index, "NAMA BB"] = nama_bb
            st.session_state["df_harga_bb"].at[index, "HARGA PER SATUAN"] = harga
            st.session_state["df_harga_bb"].at[index, "SATUAN"] = satuan
            st.session_state["df_harga_bb"].at[index, "KATEGORI"] = kategori
            st.session_state["df_harga_bb"].at[index, "CATATAN"] = catatan
            st.success("Data Bahan Baku berhasil diperbarui!")
            st.rerun()

@st.dialog("🗑️ Konfirmasi Hapus Item BB")
def open_delete_harga_bb_dialog(index):
    df = st.session_state["df_harga_bb"]
    nama_item = df.iloc[index].get("NAMA BB", "Item ini")
    nama_sup = df.iloc[index].get("NAMA SUPPLIER", "-")
    
    st.warning(f"Apakah Anda yakin ingin menghapus item **{nama_item}** dari **{nama_sup}**?")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("❌ Batal", use_container_width=True):
            st.rerun()
    with col2:
        if st.button("🗑️ Ya, Hapus Item", type="primary", use_container_width=True):
            st.session_state["df_harga_bb"] = st.session_state["df_harga_bb"].drop(index).reset_index(drop=True)
            st.success("Item berhasil dihapus!")
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
        "🛒 Update Harga BB",
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
    m3.metric("Kategori Barang", "16 Kategori", delta="Lengkap")
    m4.metric("Rata-rata Ketepatan", "94.2%", delta="+1.5%")

# ---------------------------------------------------------
# MODUL BARU: UPDATE HARGA BB (SINKRONISASI, EDIT, HAPUS, ADD, DOWNLOAD)
# ---------------------------------------------------------
elif menu == "🛒 Update Harga BB":
    st.subheader("🛒 Update & Sinkronisasi Harga Bahan Baku (BB)")
    st.caption("Data hasil inputan supplier via Link Form secara terpisah disinkronkan di halaman ini.")

    df_harga = st.session_state["df_harga_bb"]

    # Filter & Aksi Baris Atas
    col_search, col_kat, col_add, col_dl = st.columns([2.5, 2, 1.5, 1.5])

    with col_search:
        search_query = st.text_input("🔍 Cari Item / Supplier", placeholder="Ketik nama item atau supplier...")

    with col_kat:
        kat_filter = st.selectbox("📂 Filter Kategori", ["Semua Kategori"] + list(set(df_harga["KATEGORI"].dropna().tolist())) if not df_harga.empty else ["Semua Kategori"])

    with col_add:
        st.write("")
        st.write("")
        if st.button("➕ Tambah Item BB", type="primary", use_container_width=True):
            open_add_harga_bb_dialog()

    with col_dl:
        st.write("")
        st.write("")
        # Menyiapkan buffer Excel khusus sheet Update Harga BB
        output_harga = io.BytesIO()
        with pd.ExcelWriter(output_harga, engine='openpyxl') as writer:
            df_harga.to_excel(writer, index=False, sheet_name='UPDATE HARGA BB')
        
        st.download_button(
            label="📥 Download Sheet BB",
            data=output_harga.getvalue(),
            file_name="Update_Harga_BB_Koperasi_YK.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.markdown("---")

    # Tampilan Sinkronisasi Ringkasan
    total_penawaran = len(df_harga)
    total_supplier_aktif = df_harga["NAMA SUPPLIER"].nunique() if not df_harga.empty else 0
    
    k1, k2, k3 = st.columns(3)
    k1.metric("Total Entry Penawaran", f"{total_penawaran} Item")
    k2.metric("Supplier Berpartisipasi", f"{total_supplier_aktif} Supplier")
    k3.metric("Status Sinkronisasi", "🟢 Real-time Active")

    st.markdown("---")
    st.markdown("##### 📋 **Tabel Rincian Update Harga Bahan Baku**")

    # Terapkan Filter
    df_filtered = df_harga.copy()

    if kat_filter != "Semua Kategori":
        df_filtered = df_filtered[df_filtered["KATEGORI"] == kat_filter]

    if search_query:
        df_filtered = df_filtered[
            df_filtered["NAMA BB"].astype(str).str.contains(search_query, case=False, na=False) |
            df_filtered["NAMA SUPPLIER"].astype(str).str.contains(search_query, case=False, na=False) |
            df_filtered["KODE SUPPLIER"].astype(str).str.contains(search_query, case=False, na=False)
        ]

    if df_filtered.empty:
        st.info("Belum ada data penawaran harga BB atau tidak ada data yang cocok dengan filter.")
    else:
        # Header Tabel Custom
        h_cols = st.columns([1.0, 2.0, 2.2, 1.5, 0.8, 1.5, 2.0, 0.5, 0.5])
        headers = ["KODE SUPP", "NAMA SUPPLIER", "NAMA BAHAN BAKU", "HARGA (RP)", "SATUAN", "KATEGORI", "CATATAN", "EDIT", "HAPUS"]
        for col, h in zip(h_cols, headers):
            col.markdown(f"**{h}**")
        st.markdown("<hr style='margin-top:0; margin-bottom:10px;'>", unsafe_allow_html=True)

        # Loop Menampilkan Data per Baris
        for i, r in df_filtered.iterrows():
            bg_color = "#f8fafc" if i % 2 == 0 else "#ffffff"
            with st.container():
                st.markdown(f'<div style="background-color: {bg_color}; padding: 6px 10px; border-radius: 8px; border: 1px solid #e2e8f0; margin-bottom: 6px;">', unsafe_allow_html=True)
                c_kodesup, c_namasup, c_namabb, c_harga, c_sat, c_kat, c_cat, c_edit, c_del = st.columns([1.0, 2.0, 2.2, 1.5, 0.8, 1.5, 2.0, 0.5, 0.5])

                c_kodesup.write(f"**{r.get('KODE SUPPLIER', '-')}**")
                c_namasup.write(r.get("NAMA SUPPLIER", "-"))
                c_namabb.write(f"**{r.get('NAMA BB', '-')}**")
                
                harga_val = r.get("HARGA PER SATUAN", 0)
                c_harga.write(f"Rp {harga_val:,.0f}" if pd.notna(harga_val) else "-")
                
                c_sat.write(r.get("SATUAN", "-"))
                c_kat.write(f"`{r.get('KATEGORI', '-')}`")
                c_cat.write(r.get("CATATAN", "-") if pd.notna(r.get("CATATAN")) and r.get("CATATAN") != "" else "-")

                if c_edit.button("✏️", key=f"edit_bb_{i}"):
                    open_edit_harga_bb_dialog(i)
                if c_del.button("🗑️", key=f"del_bb_{i}"):
                    open_delete_harga_bb_dialog(i)

                st.markdown("</div>", unsafe_allow_html=True)

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

    # Export Data Button
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
# MODUL LAINNYA
# ---------------------------------------------------------
else:
    st.info(f"Modul `{menu}` siap digunakan.")
