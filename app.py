import streamlit as st
import pandas as pd
import numpy as np
import os

st.set_page_config(page_title="Sistem Evaluasi Supplier & Dapur Koperasi YK", layout="wide")

# Fungsi Load Data Excel / XLSB
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
                data_dict[sheet] = pd.read_excel(file_name, sheet_name=sheet)
            return data_dict, file_name
        except Exception as e:
            return None, str(e)
    return None, "File data tidak ditemukan."

data_excel, file_status = load_data()

st.title("🏭 Sistem Evaluasi Supplier & Procurement Dapur SPPG")
st.subheader("Koperasi YK")
st.caption(f"Status Data: {file_status}")
st.markdown("---")

# Sidebar Navigasi
st.sidebar.header("📌 Navigasi Modul")
menu = st.sidebar.radio(
    "Pilih Modul:",
    [
        "1. Dashboard & HET",
        "2. Kelola Data Dapur SPPG",
        "3. Penawaran WA & Update Harga PO",
        "4. Data Supplier & Matriks Jarak",
        "5. Penilaian & Penentuan Supplier per Dapur",
        "6. Analisis Komparasi Harga Pasar & HET"
    ]
)

# ---------------------------------------------------------
# MODUL 1: DASHBOARD & HET
# ---------------------------------------------------------
if menu == "1. Dashboard & HET":
    st.header("📊 Dashboard Utama & Pemantauan HET")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Dapur SPPG", "12 Dapur")
    with col2:
        st.metric("Total Supplier", "45 Supplier")
    with col3:
        st.metric("Kategori Barang", "6 Kategori")
    with col4:
        st.metric("Rata-rata Ketepatan", "94.2%")
        
    st.markdown("---")
    col_l, col_r = st.columns([2, 1])
    
    with col_l:
        st.subheader("🏆 Top 10 Supplier Terbaik")
        data_top = {
            "Peringkat": list(range(1, 11)),
            "Nama Supplier": [f"Supplier {chr(65+i)}" for i in range(10)],
            "Kategori": ["Bahan Pokok", "Daging/Ayam", "Sayur", "Bahan Pokok", "Bumbu", "Sayur", "Daging/Ayam", "Bahan Pokok", "Buah", "Bumbu"],
            "Skor Akhir": [4.85, 4.78, 4.72, 4.65, 4.60, 4.58, 4.52, 4.49, 4.45, 4.40],
            "Status": ["Sangat Direkomendasikan"] * 10
        }
        st.dataframe(pd.DataFrame(data_top), use_container_width=True)
        
    with col_r:
        st.subheader("🌐 Pemantauan Harga HET Resmi")
        st.info("Link langsung portal pemantauan HET & Bahan Pokok:")
        st.markdown("- 🔗 [Panel Harga Pangan BAPANAS](https://panelharga.badandangan.go.id/)")
        st.markdown("- 🔗 [Harga Bahan Pokok DIY (PIHPS)](https://harga.jogjaprov.go.id/)")

# ---------------------------------------------------------
# MODUL 2: KELOLA DATA DAPUR SPPG
# ---------------------------------------------------------
elif menu == "2. Kelola Data Dapur SPPG":
    st.header("🏬 Kelola Data Dapur SPPG")
    
    if data_excel and isinstance(data_excel, dict) and "Data Dapur" in data_excel:
        st.dataframe(data_excel["Data Dapur"], use_container_width=True)
    else:
        st.info("Form Penambahan Data Dapur SPPG Baru")
        with st.form("form_dapur_baru"):
            c1, c2 = st.columns(2)
            with c1:
                nama_dpr = st.text_input("Nama Dapur SPPG Baru")
                wil = st.selectbox("Wilayah", ["Bantul", "Sleman", "Gunungkidul", "Kulon Progo", "Kota Yogyakarta"])
            with c2:
                kap = st.number_input("Kapasitas (Porsi)", min_value=100, step=100)
                pj = st.text_input("Penanggung Jawab Dapur")
            if st.form_submit_button("Simpan Dapur"):
                st.success(f"Dapur {nama_dpr} berhasil ditambahkan!")

# ---------------------------------------------------------
# MODUL 3: PENAWARAN WA & UPDATE HARGA PO
# ---------------------------------------------------------
elif menu == "3. Penawaran WA & Update Harga PO":
    st.header("📲 Kirim Link Penawaran WA & Generate PO")
    
    t_wa, t_po = st.tabs(["💬 Kirim WA Penawaran Mingguan", "📄 Generate PO Supplier"])
    
    with t_wa:
        c_w1, c_w2 = st.columns(2)
        with c_w1:
            sup_nama = st.text_input("Nama Supplier", "CV Jaya Makmur")
            sup_kat = st.selectbox("Kategori Penawaran", ["Bahan Pokok", "Sayuran & Buah", "Daging & Ayam", "Bumbu"])
            no_wa = st.text_input("Nomor WA Supplier (Contoh: 628123456789)", "628123456789")
        with c_w2:
            pesan = f"Halo {sup_nama},\nMohon update penawaran harga mingguan Koperasi YK untuk kategori *{sup_kat}* via link: https://koperasi-yk.com/update-harga?sup={sup_nama.replace(' ', '%20')}\n\nTerima kasih."
            st.text_area("Pratinjau Pesan WA", pesan, height=120)
            link = f"https://wa.me/{no_wa}?text={pesan.replace(' ', '%20').replace('\n', '%0A')}"
            st.markdown(f"[👉 **Klik Disini Kirim WhatsApp**]({link})", unsafe_allow_html=True)

    with t_po:
        st.subheader("Format Purchase Order (PO) Supplier")
        st.info("Generate PO sesuai struktur data sheet harga supplier.")
        st.button("📄 Generate & Print PO")

# ---------------------------------------------------------
# MODUL 4: DATA SUPPLIER & MATRIKS JARAK
# ---------------------------------------------------------
elif menu == "4. Data Supplier & Matriks Jarak":
    st.header("🚛 Data Supplier Kompleks & Matriks Jarak")
    matriks_jarak = pd.DataFrame({
        "Nama Supplier": ["CV Jaya Makmur", "PT Sinar Pangan", "UD Berkah Mandiri"],
        "Kategori": ["Bahan Pokok", "Daging/Ayam", "Sayur"],
        "Bantul (KM)": [5.2, 12.4, 8.1],
        "Sleman (KM)": [18.0, 6.5, 14.2],
        "Kulon Progo (KM)": [25.1, 30.0, 12.0],
        "Gunungkidul (KM)": [35.0, 42.1, 38.5]
    })
    st.dataframe(matriks_jarak, use_container_width=True)

# ---------------------------------------------------------
# MODUL 5: PENILAIAN & PENENTUAN SUPPLIER PER DAPUR
# ---------------------------------------------------------
elif menu == "5. Penilaian & Penentuan Supplier per Dapur":
    st.header("🎯 Penentuan Supplier per Dapur (Scoring Multi-Kriteria)")
    
    st.subheader("📥 Input Tanda Terima Dapur & Komplain Kualitas")
    with st.form("form_skor"):
        c1, c2, c3 = st.columns(3)
        with c1:
            st.selectbox("Dapur Target", ["Dapur SPPG Bantul Utama", "Dapur SPPG Sleman Timur"])
            st.selectbox("Supplier", ["CV Jaya Makmur", "PT Sinar Pangan", "UD Berkah Mandiri"])
        with c2:
            st.time_input("Jam Terima Barang", value=pd.to_datetime("06:15").time())
            st.selectbox("Kondisi Kualitas / Komplain", ["Baik (Tanpa Komplain)", "Minor", "Sedang", "Berat", "Ditolak Total"])
        with c3:
            st.number_input("Harga Penawaran (Rp)", value=13500)
            st.number_input("Harga Target Koperasi (Rp)", value=13000)
        st.form_submit_button("Kalkulasi Rekomendasi")
        
    st.markdown("---")
    st.subheader("📋 Hasil Pemilihan Supplier Terbaik")
    res = pd.DataFrame({
        "Rekomendasi": ["Peringkat 1", "Peringkat 2"],
        "Nama Supplier": ["CV Jaya Makmur", "PT Sinar Pangan"],
        "Skor Jarak": [95, 80],
        "Skor Harga": [90, 85],
        "Skor Kualitas": [100, 90],
        "Skor Ketepatan Waktu": [90, 100],
        "TOTAL SKOR": [93.5, 88.0],
        "Keputusan": ["PILIH UTAMA", "CADANGAN"]
    })
    st.dataframe(res, use_container_width=True)

# ---------------------------------------------------------
# MODUL 6: ANALISIS KOMPARASI HARGA PASAR & HET
# ---------------------------------------------------------
elif menu == "6. Analisis Komparasi Harga Pasar & HET":
    st.header("📈 Analisis Komparasi Harga Target vs Supplier vs HET vs Pasar")
    
    comp = pd.DataFrame({
        "Nama Komoditas": ["Beras Medium (Kg)", "Minyak Goreng (Ltr)", "Daging Ayam (Kg)", "Cabai Merah (Kg)"],
        "Harga Target Koperasi": [13000, 15000, 34000, 35000],
        "Harga Supplier": [13500, 15200, 34500, 36000],
        "HET Bantul": [13100, 15000, 35000, 37000],
        "HET Sleman": [13200, 15000, 35000, 36500],
        "HET Gunungkidul": [13300, 15500, 36000, 38000],
        "HET Kulon Progo": [13100, 15000, 34800, 36000],
        "Survey Pasar": [13400, 15300, 35500, 37500]
    })
    st.dataframe(comp, use_container_width=True)
    st.line_chart(comp.set_index("Nama Komoditas")[["Harga Target Koperasi", "Harga Supplier", "Survey Pasar"]])
