import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import random
import string
import urllib.parse

# ---------------------------------------------------------
# 1. KONFIGURASI HALAMAN
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

# CSS Header & Style
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
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 2. INITIALIZE SESSION STATE
# ---------------------------------------------------------
if "df_dapur_state" not in st.session_state:
    st.session_state["df_dapur_state"] = pd.DataFrame({
        "KODE": ["PAKEM", "NGPLK", "SLMN4"],
        "NAMA DAPUR": ["PAKEM (Hargobinangun)", "NGEMPLAK (Umbulmartani 1)", "SLEMAN 4 (Triharjo)"],
        "ALAMAT": ["Jl. Kaliurang Km 22", "Jl. Kaliurang Km 15.5", "Jl. Letkol Subadri"],
        "PIC": ["SHINTA", "SHINTA", "CAHYO"],
        "LATITUDE": [-7.618382, -7.675789, -7.700475],
        "LONGITUDE": [110.426078, 110.417100, 110.342646],
        "KOTA/KABUPATEN": ["SLEMAN", "SLEMAN", "SLEMAN"]
    })

if "df_supplier_state" not in st.session_state:
    st.session_state["df_supplier_state"] = pd.DataFrame([
        {
            "KODE SUPPLIER": "SUP-001",
            "NAMA SUPPLIER": "UD Tani Makmur",
            "NO WA": "6281234567890",
            "JENIS BB 1": "Ayam",
            "JENIS BB 2": "Telur Ayam",
            "JENIS BB 3": "-",
            "TOKEN": buat_token(),
            "LINK FORM": f"{WEB_APP_URL}?token=sample1",
            "PIC": "Budi",
            "ALAMAT": "Sleman"
        }
    ])

if "df_harga_bb" not in st.session_state:
    st.session_state["df_harga_bb"] = pd.DataFrame(columns=[
        "KODE SUPPLIER", "NAMA SUPPLIER", "NAMA BB", "HARGA PER SATUAN", "SATUAN", "KATEGORI", "CATATAN"
    ])

# ---------------------------------------------------------
# 3. HANDLER UNTUK FORM EXTERNAL VENDOR (MENGGUNAKAN QUERY PARAMS)
# ---------------------------------------------------------
query_params = st.query_params
if "token" in query_params:
    token_diterima = query_params["token"]
    kategori_diterima = query_params.get("kat", "ALL")
    
    st.title("📝 Form Penawaran Harga Supplier")
    st.caption("Koperasi YK — Sistem Informasi Procurement & SPPG")
    st.markdown("---")

    list_kat_supplier = [k.replace("_", " ") for k in kategori_diterima.split(",")] if kategori_diterima != "ALL" else ["Semua Kategori"]

    with st.form("form_input_harga_supplier"):
        st.subheader("🛒 Input Penawaran Harga")
        nama_barang = st.text_input("Nama Bahan Baku / Barang", placeholder="Contoh: Telur Ayam Ras")
        c1, c2 = st.columns(2)
        with c1:
            harga_penawaran = st.number_input("Harga Penawaran (Rp)", min_value=0, step=500, value=10000)
            satuan = st.selectbox("Satuan", ["Kg", "Liter", "Ikat", "Pcs", "Karton", "Ekor", "Pack"])
        with c2:
            kategori_pilihan = st.selectbox("Kategori Barang", options=list_kat_supplier)
            catatan = st.text_input("Catatan / Merek (Opsional)")

        if st.form_submit_button("🚀 Kirim Penawaran Harga", use_container_width=True):
            new_entry = pd.DataFrame([{
                "KODE SUPPLIER": "EXTERNAL",
                "NAMA SUPPLIER": "Vendor Form",
                "NAMA BB": nama_barang,
                "HARGA PER SATUAN": harga_penawaran,
                "SATUAN": satuan,
                "KATEGORI": kategori_pilihan,
                "CATATAN": catatan
            }])
            st.session_state["df_harga_bb"] = pd.concat([st.session_state["df_harga_bb"], new_entry], ignore_index=True)
            st.success("Penawaran harga berhasil dikirim!")

    # PERHATIKAN: st.stop() WAJIB HANYA DI DALAM BLOK TOKEN INI
    st.stop()

# ---------------------------------------------------------
# 4. SIDEBAR NAVIGASI UTAMA
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
    
    # Pilih modul
    menu = st.radio("Pilih Modul:", modul_options, index=3)
    st.markdown("---")
    st.caption("Status Data: 🟢 Sistem Evaluasi Supplier & Dapur.xlsb")

# HEADER UTAMA
st.markdown("""
<div class="header-card">
    <div>
        <h2 style="margin:0; color: white;">🏭 Enterprise Procurement System</h2>
        <p style="margin:0; opacity: 0.8;">Koperasi YK — Sistem Evaluasi Supplier & Dapur SPPG</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 5. RENDER MODUL BERDASARKAN PILIHAN MENU
# ---------------------------------------------------------
if menu == "📊 Dashboard & HET":
    st.subheader("📊 Dashboard Utama & Pemantauan HET")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Dapur SPPG", len(st.session_state["df_dapur_state"]))
    m2.metric("Total Supplier", len(st.session_state["df_supplier_state"]))
    m3.metric("Kategori Barang", "16 Kategori")
    m4.metric("Rata-rata Ketepatan", "94.2%")

elif menu == "🏬 Kelola Data Dapur":
    st.subheader("🏬 Manajemen Data Dapur SPPG & Peta Sebaran")
    st.dataframe(st.session_state["df_dapur_state"], use_container_width=True)

elif menu == "🤝 Data Supplier & Link Form":
    st.subheader("🤝 Kelola Data Supplier & Link Web App")
    st.dataframe(st.session_state["df_supplier_state"], use_container_width=True)

elif menu == "💬 WA & PO Generator":
    st.subheader("💬 WhatsApp Message & PO Link Generator")
    st.info("Modul WA & PO Generator siap digunakan.")
    
    df_sup = st.session_state["df_supplier_state"]
    list_supplier_display = [f"{r['KODE SUPPLIER']} - {r['NAMA SUPPLIER']}" for _, r in df_sup.iterrows()]
    selected_sup = st.selectbox("Pilih Supplier Target:", list_supplier_display)
    
    idx_sup = list_supplier_display.index(selected_sup)
    row_sup = df_sup.iloc[idx_sup]
    
    LIST_KATEGORI = ["Ayam", "Beras", "Buah", "Daging", "Ikan", "Sembako", "Telur Ayam"]
    selected_kategori = st.multiselect("Pilih Kategori Barang:", options=LIST_KATEGORI, default=["Ayam"])
    
    token_sup = row_sup.get("TOKEN", "sample_token")
    kat_param = ",".join([k.replace(" ", "_") for k in selected_kategori])
    link_penawaran = f"{WEB_APP_URL}?token={token_sup}&kat={kat_param}"
    
    pesan_wa = f"""Yth. Bapak/Ibu Vendor *{row_sup.get('NAMA SUPPLIER')}*,

Mohon bantuan untuk memperbarui / mengirimkan penawaran harga bahan baku kategori:
*{', '.join(selected_kategori)}*

Silakan isi via link resmi berikut:
🔗 {link_penawaran}

Terima kasih.
*Tim Procurement Koperasi YK*"""

    st.text_area("Draft Pesan WA:", value=pesan_wa, height=180)
    
    no_wa = str(row_sup.get("NO WA", "")).strip()
    if no_wa:
        wa_url = f"https://wa.me/{no_wa}?text={urllib.parse.quote(pesan_wa)}"
        st.markdown(f'<a href="{wa_url}" target="_blank"><button style="background-color:#25D366; color:white; border:none; padding:10px; border-radius:6px; cursor:pointer; width:100%;">💬 Kirim via WhatsApp Direct</button></a>', unsafe_allow_html=True)

elif menu == "🚛 Matriks Jarak":
    st.subheader("🚛 Matriks Jarak & Rute Distribusi Dapur SPPG")

elif menu == "🎯 Scoring & Evaluasi":
    st.subheader("🎯 Scoring & Evaluasi Kinerja Supplier")

elif menu == "📈 HET & Komparasi Pasar":
    st.subheader("📈 Pemantauan Harga HET & Komparasi Pasar")
