import streamlit as st
import pandas as pd
import numpy as np
import os
import io
import urllib.parse

# Konfigurasi Halaman Streamlit
st.set_page_config(
    page_title="SPPG Procurement Engine | Koperasi YK",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling CSS
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
    .stSidebar [data-testid="stRadioButton"] > label {
        display: none;
    }
</style>
""", unsafe_allow_html=True)

# --- LOAD DATA DARI EXCEL ---
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

# 1. Inisialisasi State Data Dapur
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
        cols_to_keep = [c for c in ["KODE", "NAMA DAPUR", "ALAMAT", "PIC", "LATITUDE", "LONGITUDE", "KOTA/KABUPATEN"] if c in raw_df.columns]
        st.session_state["df_dapur_state"] = raw_df[cols_to_keep].reset_index(drop=True)
    else:
        st.session_state["df_dapur_state"] = pd.DataFrame({
            "KODE": ["PAKEM", "NGPLK"],
            "NAMA DAPUR": ["PAKEM (Hargobinangun)", "NGEMPLAK (Umbulmartani 1)"],
            "ALAMAT": ["Jl. Kaliurang Km 22", "Jl. Kaliurang Km 15.5"],
            "PIC": ["SHINTA", "SHINTA"],
            "LATITUDE": [-7.618382, -7.675789],
            "LONGITUDE": [110.426078, 110.417100],
            "KOTA/KABUPATEN": ["SLEMAN", "SLEMAN"]
        })

# 2. Inisialisasi State Data Supplier
if "df_supplier_state" not in st.session_state:
    if data_excel and isinstance(data_excel, dict) and "Data Supplier" in data_excel:
        st.session_state["df_supplier_state"] = data_excel["Data Supplier"].dropna(how="all").reset_index(drop=True)
    else:
        st.session_state["df_supplier_state"] = pd.DataFrame({
            "KODE SUPPLIER": ["SUP-001", "SUP-002"],
            "NAMA SUPPLIER": ["UD Berkah Tani", "CV Sayur Segar"],
            "KATEGORI": ["Sayur & Buah", "Daging & Ayam"],
            "NO WA": ["6285169796974", "6281234567890"],
            "ALAMAT": ["Sleman, Yogyakarta", "Bantul, Yogyakarta"],
            "STATUS": ["Aktif", "Aktif"]
        })

# 3. Inisialisasi Format Bahan Baku (Sheet Harga Supplier)
if "df_bahan_baku" not in st.session_state:
    if data_excel and isinstance(data_excel, dict) and "Harga Supplier" in data_excel:
        st.session_state["df_bahan_baku"] = data_excel["Harga Supplier"].copy()
    else:
        st.session_state["df_bahan_baku"] = pd.DataFrame({
            "NO": [1, 2, 3, 4, 5],
            "NAMA BAHAN BAKU": ["Beras Medium", "Daging Ayam Ras", "Telur Ayam Ras", "Minyak Goreng", "Bawang Merah"],
            "SATUAN": ["Kg", "Kg", "Kg", "Liter", "Kg"]
        })

# --- POP-UP DIALOG UNTUK SUPPLIER ---
@st.dialog("➕ Tambah Data Supplier Baru")
def open_add_supplier_dialog():
    with st.form("form_add_sup"):
        c1, c2 = st.columns(2)
        with c1:
            kode = st.text_input("Kode Supplier (Contoh: SUP-003)")
            nama = st.text_input("Nama Supplier / Toko")
            kategori = st.selectbox("Kategori", ["Sayur & Buah", "Daging & Ayam", "Sembako", "Bumbu", "Lainnya"])
        with c2:
            no_wa = st.text_input("Nomor WA (Gunakan format 62..., Contoh: 6285169796974)", value="6285169796974")
            alamat = st.text_area("Alamat Lengkap", height=70)
            status = st.selectbox("Status Supplier", ["Aktif", "Non-Aktif"])
            
        if st.form_submit_button("✨ Simpan Supplier", use_container_width=True):
            new_row = pd.DataFrame([{
                "KODE SUPPLIER": kode,
                "NAMA SUPPLIER": nama,
                "KATEGORI": kategori,
                "NO WA": no_wa,
                "ALAMAT": alamat,
                "STATUS": status
            }])
            st.session_state["df_supplier_state"] = pd.concat([st.session_state["df_supplier_state"], new_row], ignore_index=True)
            st.success("Supplier baru berhasil disimpan!")
            st.rerun()

@st.dialog("🗑️ Hapus Supplier")
def open_del_supplier_dialog(idx):
    df = st.session_state["df_supplier_state"]
    nama = df.iloc[idx].get("NAMA SUPPLIER", "Supplier")
    st.warning(f"Apakah Anda yakin ingin menghapus **{nama}**?")
    c1, c2 = st.columns(2)
    if c1.button("❌ Batal", use_container_width=True):
        st.rerun()
    if c2.button("🗑️ Ya, Hapus", type="primary", use_container_width=True):
        st.session_state["df_supplier_state"] = df.drop(idx).reset_index(drop=True)
        st.success("Data supplier berhasil dihapus!")
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
        "🤝 Data Supplier",
        "💬 WA Generator",
        "🚛 Matriks Jarak",
        "🎯 Scoring & Evaluasi",
        "📈 HET & Komparasi Pasar"
    ]
    
    menu = st.radio("Pilih Modul:", modul_options, index=3)
    st.markdown("---")
    st.caption(f"Status File: 🟢 {file_status}")

# --- HEADER UTAMA ---
st.markdown("""
<div class="header-card">
    <h2 style="margin:0; color: white;">🏭 Enterprise Procurement System</h2>
    <p style="margin:0; opacity: 0.8;">Koperasi YK — Sistem Evaluasi Supplier & Dapur SPPG</p>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# MODUL: DATA SUPPLIER
# ---------------------------------------------------------
if menu == "🤝 Data Supplier":
    st.subheader("🤝 Kelola Data Supplier SPPG")
    
    col_a, col_b = st.columns([1, 2])
    with col_a:
        if st.button("➕ Tambah Supplier Baru", type="primary", use_container_width=True):
            open_add_supplier_dialog()
            
    with col_b:
        # Export ke Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            st.session_state["df_supplier_state"].to_excel(writer, index=False, sheet_name='Data Supplier')
            st.session_state["df_bahan_baku"].to_excel(writer, index=False, sheet_name='Harga Supplier')
        
        st.download_button(
            label="📥 Download Data & Format Harga (Excel)",
            data=output.getvalue(),
            file_name="Data_Supplier_dan_Format_Harga.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    st.markdown("---")
    
    df_sup = st.session_state["df_supplier_state"]
    if not df_sup.empty:
        st.dataframe(df_sup, use_container_width=True)
        st.markdown("##### ⚙️ **Aksi Hapus Data Supplier**")
        sup_to_del = st.selectbox("Pilih Supplier yang ingin dihapus:", options=range(len(df_sup)), format_func=lambda x: f"{df_sup.iloc[x].get('KODE SUPPLIER', '')} - {df_sup.iloc[x].get('NAMA SUPPLIER', '')}")
        if st.button("🗑️ Hapus Supplier Terpilih", type="primary"):
            open_del_supplier_dialog(sup_to_del)
    else:
        st.info("Belum ada data supplier.")

# ---------------------------------------------------------
# MODUL: WA GENERATOR (TANPA PO, DIRECT SEND TO WA)
# ---------------------------------------------------------
elif menu == "💬 WA Generator":
    st.subheader("💬 Request Update Harga Supplier via WhatsApp")
    
    df_sup = st.session_state["df_supplier_state"]
    df_bahan = st.session_state["df_bahan_baku"]
    
    col_left, col_right = st.columns([1, 1])
    
    with col_left:
        st.markdown("##### 📱 **1. Pengaturan Pengirim & Penerima**")
        sender_wa = st.text_input("Nomor Pengirim (Koperasi YK):", value="085169796974")
        
        if not df_sup.empty:
            sup_options = df_sup["NAMA SUPPLIER"].tolist()
            selected_sup_name = st.selectbox("Pilih Supplier Tujuan:", sup_options)
            
            sup_data = df_sup[df_sup["NAMA SUPPLIER"] == selected_sup_name].iloc[0]
            target_wa = str(sup_data.get("NO WA", "6285169796974")).replace("-", "").replace(" ", "").replace("+", "")
            if target_wa.startswith("0"):
                target_wa = "62" + target_wa[1:]
        else:
            selected_sup_name = "Supplier Utama"
            target_wa = "6285169796974"
            
        st.info(f"Target Kirim ke WA: **{target_wa}** ({selected_sup_name})")
        
        # Generator Draf Format Bahan Baku dari Sheet Harga Supplier
        draf_list_bahan = "FORMAT UPDATE HARGA BAHAN BAKU:\n"
        if "NAMA BAHAN BAKU" in df_bahan.columns and "SATUAN" in df_bahan.columns:
            for idx, row in df_bahan.iterrows():
                nama_b = row["NAMA BAHAN BAKU"]
                sat = row["SATUAN"]
                draf_list_bahan += f"{idx+1}. {nama_b} ({sat}): Rp ...\n"
        else:
            draf_list_bahan += "1. Beras Medium (Kg): Rp ...\n2. Daging Ayam Ras (Kg): Rp ...\n"

        default_message = f"""Yth. {selected_sup_name},

Kami dari Koperasi YK (SPPG Procurement) memohon informasi update harga bahan baku terbaru untuk periode minggu ini.

Berikut daftar bahan baku yang kami minta penawarannya:

{draf_list_bahan}
Mohon balaskan pesan ini dengan menyertakan harga update Anda. Terima kasih banyak atas kerjasamanya.

Regards,
Procurement Team Koperasi YK
WA Admin: {sender_wa}"""

    with col_right:
        st.markdown("##### ✏️ **2. Edit & Sesuaikan Teks Pesan WA**")
        custom_message = st.text_area("Format Pesan (Bisa Diedit Sesuai Kebutuhan):", value=default_message, height=350)
        
        encoded_message = urllib.parse.quote(custom_message)
        wa_link = f"https://wa.me/{target_wa}?text={encoded_message}"
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(f"""
        <a href="{wa_link}" target="_blank" style="text-decoration:none;">
            <div style="background-color: #25D366; color: white; padding: 14px; border-radius: 10px; text-align: center; font-weight: bold; font-size: 18px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                📲 KLIK SEKALI UNTUK KIRIM WHATSAPP
            </div>
        </a>
        """, unsafe_allow_html=True)

# ---------------------------------------------------------
# MODUL LAINNYA
# ---------------------------------------------------------
elif menu == "📊 Dashboard & HET":
    st.subheader("📊 Dashboard Utama & Pemantauan HET")
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Dapur SPPG", len(st.session_state["df_dapur_state"]))
    m2.metric("Total Supplier", len(st.session_state["df_supplier_state"]))
    m3.metric("Daftar Item Bahan Baku", len(st.session_state["df_bahan_baku"]))

else:
    st.subheader(f"Modul {menu}")
    st.info("Modul ini siap dikembangkan sesuai kebutuhan.")
