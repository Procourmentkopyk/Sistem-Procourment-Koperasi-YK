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

# --- CUSTOM CSS ---
st.markdown("""
<style>
    .header-card {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 10px rgba(0,0,0,0.15);
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

# --- LOAD DATA EXCEL MASTER ---
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

# --- 1. SESSION STATE: DATA SUPPLIER & LINK ---
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

# --- 2. SESSION STATE: HARGA BB MASTER ---
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

# --- 3. SESSION STATE: TEMPLATE PESAN WA MINGGUAN ---
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

# --- POP-UP DIALOG MANAGEMENT ---
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

# --- SIDEBAR NAVIGASI ---
with st.sidebar:
    st.markdown("### ⚡ **SPPG Engine**")
    st.caption("Koperasi YK")
    st.markdown("---")
    
    modul_options = [
        "💬 WA Generator & Broadcast",
        "🤝 Data Supplier & Link",
        "📊 Master Harga Bahan Baku"
    ]
    
    menu = st.radio("Pilih Modul:", modul_options, index=0)
    st.markdown("---")
    st.caption(f"Status Excel: 🟢 {file_status}")

# --- HEADER UTAMA ---
st.markdown("""
<div class="header-card">
    <h2 style="margin:0; color: white;">🏭 Enterprise Procurement System</h2>
    <p style="margin:0; opacity: 0.8;">Koperasi YK — Sistem Evaluasi Supplier & Request Harga WA</p>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# MODUL 1: WA GENERATOR & BROADCAST (1-KLIK)
# ---------------------------------------------------------
if menu == "💬 WA Generator & Broadcast":
    st.subheader("💬 Broadcast & Request Update Harga via WhatsApp")
    
    df_sup = st.session_state["df_supplier_state"]
    df_bb = st.session_state["df_harga_bb"]
    cfg = st.session_state["wa_template_config"]
    
    # --- PENGATURAN TEMPLATE PESAN MINGGUAN ---
    with st.expander("⚙️ **Atur Template Pesan & Periode Mingguan (Klik untuk Mengubah)**", expanded=False):
        c_tp1, c_tp2 = st.columns(2)
        with c_tp1:
            st.session_state["wa_template_config"]["periode_default"] = st.text_input(
                "Periode Pengiriman / Minggu Ke:", 
                value=cfg["periode_default"]
            )
            st.session_state["wa_template_config"]["header_admin"] = st.text_input(
                "Nama Pengirim / Divisi:", 
                value=cfg["header_admin"]
            )
        with c_tp2:
            st.session_state["wa_template_config"]["wa_admin"] = st.text_input(
                "Nomor WA Admin:", 
                value=cfg["wa_admin"]
            )
            st.caption("💡 Tag Otomatis yang tersedia: `{NAMA_SUPPLIER}`, `{PERIODE_MINGGU}`, `{LINK_FORM}`, `{DAFTAR_BAHAN}`, `{HEADER_ADMIN}`, `{WA_ADMIN}`")
            
        st.session_state["wa_template_config"]["template_text"] = st.text_area(
            "Format Pesan Master (Akan diterapkan ke semua supplier):",
            value=cfg["template_text"],
            height=200
        )
        st.success(" Template pesan tersimpan otomatis untuk sesi ini.")

    st.markdown("---")
    
    # --- PILIH SUPPLIER & GENERATE PESAN ---
    col_left, col_right = st.columns([1, 1])
    
    with col_left:
        st.markdown("##### 📱 **1. Pilih Supplier Tujuan**")
        
        if not df_sup.empty:
            sup_names = df_sup["NAMA SUPPLIER"].tolist()
            selected_sup_name = st.selectbox("Pilih Supplier:", sup_names)
            
            sup_row = df_sup[df_sup["NAMA SUPPLIER"] == selected_sup_name].iloc[0]
            kode_sup = str(sup_row.get("KODE SUPPLIER", ""))
            token_sup = str(sup_row.get("TOKEN", ""))
            
            # Format Nomor WA
            target_wa = str(sup_row.get("NO WA", "6285169796974")).replace("-", "").replace(" ", "").replace("+", "")
            if target_wa.startswith("0"):
                target_wa = "62" + target_wa[1:]
                
            form_link = f"{WEB_APP_URL}?token={token_sup}"
            
            # Ambil jenis BB supplier
            jenis_bb_list = []
            for col in sup_row.index:
                if "JENIS" in str(col).upper() and pd.notna(sup_row[col]) and str(sup_row[col]).strip() != "":
                    jenis_bb_list.append(normalisasi(sup_row[col]))
        else:
            selected_sup_name = "Supplier Tani"
            kode_sup = "SUP-001"
            target_wa = "6285169796974"
            form_link = WEB_APP_URL
            jenis_bb_list = []

        # Filter Daftar Bahan Baku Sesuai Jenis BB Supplier (Logic Apps Script)
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

        # Menyusun Pesan Akhir Berdasarkan Template
        pesan_jadi = cfg["template_text"].format(
            NAMA_SUPPLIER=selected_sup_name,
            PERIODE_MINGGU=cfg["periode_default"],
            LINK_FORM=form_link,
            DAFTAR_BAHAN=draf_list_bahan,
            HEADER_ADMIN=cfg["header_admin"],
            WA_ADMIN=cfg["wa_admin"]
        )

        st.info(f"📌 **Target**: {selected_sup_name} | **No WA**: +{target_wa}\n\n📋 **Total Item Relevan**: {item_count} Bahan Baku")

    with col_right:
        st.markdown("##### ✏️ **2. Pratinjau & Edit Pesan (Khusus Supplier Ini)**")
        
        # Pesan fleksibel yang dapat disesuaikan lagi per supplier jika ada catatan khusus
        final_message = st.text_area("Pesan WhatsApp Siap Kirim:", value=pesan_jadi, height=320)
        
        encoded_message = urllib.parse.quote(final_message)
        wa_link = f"https://wa.me/{target_wa}?text={encoded_message}"
        
        st.markdown(f"""
        <div class="wa-card">
            <p style="margin:0; font-weight:bold; color:#1e293b;">🚀 Siap Kirim via WhatsApp</p>
            <p style="margin:5px 0 0 0; font-size:13px; color:#64748b;">Klik tombol di bawah untuk membuka aplikasi WhatsApp Web / Mobile dan langsung mengirimkan pesan di atas.</p>
            <a href="{wa_link}" target="_blank" class="wa-button">
                📲 KIRIM PESAN WA (1-KLIK)
            </a>
        </div>
        """, unsafe_allow_html=True)

# ---------------------------------------------------------
# MODUL 2: DATA SUPPLIER & LINK
# ---------------------------------------------------------
elif menu == "🤝 Data Supplier & Link":
    st.subheader("🤝 Kelola Data Supplier & Token Form Web App")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        if st.button("➕ Tambah Supplier Baru", type="primary", use_container_width=True):
            open_add_supplier_dialog()
            
    with col2:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            st.session_state["df_harga_bb"].to_excel(writer, index=False, sheet_name='UPDATE HARGA BB')
            st.session_state["df_supplier_state"].to_excel(writer, index=False, sheet_name='Supplier Link')
        
        st.download_button(
            label="📥 Download Data Supplier & Master Harga (Excel)",
            data=output.getvalue(),
            file_name="Data_Supplier_dan_Master_Harga.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    st.markdown("---")
    df_sup = st.session_state["df_supplier_state"]
    if not df_sup.empty:
        df_sup["LINK FORM"] = df_sup["TOKEN"].apply(lambda t: f"{WEB_APP_URL}?token={t}" if pd.notna(t) else "")
        st.dataframe(df_sup, use_container_width=True)
    else:
        st.info("Belum ada data supplier.")

# ---------------------------------------------------------
# MODUL 3: MASTER HARGA BAHAN BAKU
# ---------------------------------------------------------
elif menu == "📊 Master Harga Bahan Baku":
    st.subheader("📊 Master Data Bahan Baku (`UPDATE HARGA BB`)")
    st.dataframe(st.session_state["df_harga_bb"], use_container_width=True)
