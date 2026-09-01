import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Sistem Evaluasi Supplier & Dapur Koperasi YK", layout="wide")

# Title & Header
st.title("🏭 Sistem Evaluasi Supplier & Procurement Dapur SPPG")
st.subheader("Koperasi YK")
st.markdown("---")

# Sidebar Menu
st.sidebar.header("📌 Navigasi Menu")
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
    
    # Sample Stats
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Dapur SPPG", "12 Dapur")
    with col2:
        st.metric("Total Supplier Terdaftar", "45 Supplier")
    with col3:
        st.metric("Kategori Barang", "6 Kategori")
    with col4:
        st.metric("Rata-rata Ketepatan", "94.2%")
        
    st.markdown("---")
    
    col_left, col_right = st.columns([2, 1])
    
    with col_left:
        st.subheader("🏆 Top 10 Supplier Terbaik")
        data_top = {
            "Peringkat": list(range(1, 11)),
            "Nama Supplier": [f"Supplier {chr(65+i)}" for i in range(10)],
            "Kategori": ["Bahan Pokok", "Daging/Ayam", "Sayur", "Bahan Pokok", "Bumbu", "Sayur", "Daging/Ayam", "Bahan Pokok", "Buah", "Bumbu"],
            "Skor Akhir": [4.85, 4.78, 4.72, 4.65, 4.60, 4.58, 4.52, 4.49, 4.45, 4.40],
            "Status": ["Sangat Direkomendasikan"] * 10
        }
        st.dataframe(pd.DataFrame(data_top), use_container_width=True)
        
    with col_right:
        st.subheader("🌐 Web Pemantauan HET")
        st.info("Akses langsung ke portal resmi pemantauan harga pangan & HET daerah DIY:")
        st.markdown("- 🔗 [Panel Harga Pangan Badan Pangan Nasional](https://panelharga.badandangan.go.id/)")
        st.markdown("- 🔗 [Sistem Informasi Harga Bahan Pokok DIY (PIHPS)](https://harga.jogjaprov.go.id/)")
        st.markdown("- 🔗 [Data HET Kabupaten Bantul & Sleman](#)")

# ---------------------------------------------------------
# MODUL 2: KELOLA DATA DAPUR SPPG
# ---------------------------------------------------------
elif menu == "2. Kelola Data Dapur SPPG":
    st.header("🏬 Pengelolaan Data Dapur SPPG")
    
    tab1, tab2 = st.tabs(["📋 Daftar Dapur SPPG", "➕ Tambah / Edit Dapur Baru"])
    
    with tab1:
        data_dapur = {
            "ID Dapur": ["DPR-01", "DPR-02", "DPR-03", "DPR-04"],
            "Nama Dapur": ["Dapur SPPG Bantul Utama", "Dapur SPPG Sleman Timur", "Dapur SPPG Kulon Progo", "Dapur SPPG Gunungkidul"],
            "Wilayah/Kabupaten": ["Bantul", "Sleman", "Kulon Progo", "Gunungkidul"],
            "Kapasitas (Porsi)": [1500, 2000, 1200, 1000],
            "Penanggung Jawab": ["Pak Budi", "Bu Siti", "Pak Ahmad", "Bu Rina"]
        }
        st.dataframe(pd.DataFrame(data_dapur), use_container_width=True)
        
    with tab2:
        st.subheader("Form Input / Edit Dapur SPPG")
        with st.form("form_dapur"):
            col_a, col_b = st.columns(2)
            with col_a:
                id_dapur = st.text_input("ID Dapur", value="DPR-05")
                nama_dapur = st.text_input("Nama Dapur SPPG")
                wilayah = st.selectbox("Wilayah (Kabupaten)", ["Bantul", "Sleman", "Gunungkidul", "Kulon Progo", "Kota Yogyakarta"])
            with col_b:
                kapasitas = st.number_input("Kapasitas Porsi per Hari", min_value=100, step=100)
                pj = st.text_input("Nama Penanggung Jawab Dapur")
                kontak = st.text_input("No. WhatsApp Dapur")
                
            btn_simpan_dapur = st.form_submit_button("Simpan Data Dapur")
            if btn_simpan_dapur:
                st.success(f"Data {nama_dapur} berhasil disimpan!")

# ---------------------------------------------------------
# MODUL 3: PENAWARAN WA & UPDATE HARGA PO
# ---------------------------------------------------------
elif menu == "3. Penawaran WA & Update Harga PO":
    st.header("📲 Kirim Link Penawaran WA & Pembuatan PO")
    
    tab_wa, tab_po = st.tabs(["💬 Send Link Penawaran WhatsApp", "📄 Generate Purchase Order (PO)"])
    
    with tab_wa:
        st.subheader("Kirim Pesan Pengisian Harga Mingguan via WA")
        col_w1, col_w2 = st.columns(2)
        with col_w1:
            sup_nama = st.selectbox("Pilih Supplier", ["CV Jaya Makmur", "PT Sinar Pangan", "UD Berkah Mandiri"])
            sup_kat = st.selectbox("Kategori Penawaran", ["Bahan Pokok", "Sayuran & Buah", "Daging & Ayam", "Bumbu"])
            no_wa = st.text_input("Nomor WA Supplier (Format: 628xxx)", "628123456789")
        with col_w2:
            pesan_wa = f"Halo {sup_nama},\nMohon update penawaran harga mingguan Koperasi YK untuk kategori *{sup_kat}* melalui link form berikut: https://koperasi-yk.com/update-harga?sup={sup_nama.replace(' ', '%20')}\n\nTerima kasih."
            st.text_area("Pratinjau Pesan WA", pesan_wa, height=120)
            
            link_wa = f"https://wa.me/{no_wa}?text={pesan_wa.replace(' ', '%20').replace('\n', '%0A')}"
            st.markdown(f"[👉 **Klik Disini untuk Kirim via WhatsApp**]({link_wa})", unsafe_allow_html=True)
            
    with tab_po:
        st.subheader("Pembuatan Purchase Order (PO) Supplier")
        st.info("Format disesuaikan otomatis dengan sheet harga supplier.")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            po_supplier = st.selectbox("Pilih Supplier PO", ["CV Jaya Makmur", "PT Sinar Pangan"])
            po_dapur = st.selectbox("Target Dapur Tujuan", ["Dapur SPPG Bantul Utama", "Dapur SPPG Sleman Timur"])
        with col_p2:
            tgl_po = st.date_input("Tanggal Pengiriman")
            
        st.subheader("Rincian Item Barang")
        po_items = pd.DataFrame({
            "Nama Barang": ["Bras Medium", "Minyak Goreng", "Gula Pasir"],
            "Satuan": ["Kg", "Liter", "Kg"],
            "Volume Target": [500, 200, 100],
            "Harga Sepakat (Rp)": [13500, 15500, 17000],
            "Total (Rp)": [6750000, 3100000, 1700000]
        })
        st.dataframe(po_items, use_container_width=True)
        st.button("📄 Cetak / Download PO")

# ---------------------------------------------------------
# MODUL 4: DATA SUPPLIER & MATRIKS JARAK
# ---------------------------------------------------------
elif menu == "4. Data Supplier & Matriks Jarak":
    st.header("🚛 Data Supplier Kompleks & Matriks Jarak ke Dapur")
    
    st.subheader("📍 Matriks Jarak Supplier ke Dapur SPPG (dalam KM)")
    matriks_jarak = pd.DataFrame({
        "Nama Supplier": ["CV Jaya Makmur", "PT Sinar Pangan", "UD Berkah Mandiri", "Toko Sumber Rejeki"],
        "Kategori": ["Bahan Pokok", "Daging/Ayam", "Sayur", "Bumbu"],
        "Dapur Bantul (KM)": [5.2, 12.4, 8.1, 15.0],
        "Dapur Sleman (KM)": [18.0, 6.5, 14.2, 9.8],
        "Dapur Kulon Progo (KM)": [25.1, 30.0, 12.0, 22.4],
        "Dapur Gunungkidul (KM)": [35.0, 42.1, 38.5, 28.0]
    })
    st.dataframe(matriks_jarak, use_container_width=True)

# ---------------------------------------------------------
# MODUL 5: PENILAIAN & PENENTUAN SUPPLIER PER DAPUR
# ---------------------------------------------------------
elif menu == "5. Penilaian & Penentuan Supplier per Dapur":
    st.header("🎯 System Evaluasi & Skor Supplier per Dapur")
    
    col_sel1, col_sel2 = st.columns(2)
    with col_sel1:
        dapur_tujuan = st.selectbox("Pilih Dapur Target", ["Dapur SPPG Bantul Utama", "Dapur SPPG Sleman Timur"])
    with col_sel2:
        kategori_target = st.selectbox("Filter Kategori Barang", ["Bahan Pokok", "Daging/Ayam", "Sayur", "Bumbu"])
        
    st.markdown("---")
    st.subheader("📥 Form Input Penerimaan Barang & Komplain Kualitas (Operational Input)")
    
    with st.form("form_penerimaan"):
        c1, c2, c3 = st.columns(3)
        with c1:
            input_sup = st.selectbox("Supplier", ["CV Jaya Makmur", "PT Sinar Pangan", "UD Berkah Mandiri"])
            target_jam = st.time_input("Target Jam Tiba", value=pd.to_datetime("06:00").time())
            aktual_jam = st.time_input("Jam Aktual Tiba Dapur", value=pd.to_datetime("06:15").time())
        with c2:
            status_kualitas = st.selectbox("Kondisi Kualitas Barang", ["Baik (Tanpa Komplain)", "Minor", "Sedang", "Berat", "Ditolak Total"])
            harga_penawaran = st.number_input("Harga Penawaran Supplier (Rp)", value=13500)
            harga_target_kop = st.number_input("Target Harga Koperasi (Rp)", value=13000)
        with c3:
            jarak_km = st.number_input("Jarak ke Dapur (KM)", value=5.2)
            catatan_terima = st.text_area("Catatan Tanda Terima")
            
        btn_eval = st.form_submit_button("Kalkulasi Skor Evaluasi")
        
    st.markdown("---")
    st.subheader("📋 Hasil Ranking Rekomendasi Supplier untuk " + dapur_tujuan)
    
    res_ranking = pd.DataFrame({
        "Rekomendasi": ["Peringkat 1", "Peringkat 2", "Peringkat 3"],
        "Nama Supplier": ["CV Jaya Makmur", "PT Sinar Pangan", "UD Berkah Mandiri"],
        "Skor Jarak": [95, 80, 70],
        "Skor Harga": [90, 85, 75],
        "Skor Kualitas": [100, 90, 60],
        "Skor Waktu Tiba": [90, 100, 80],
        "TOTAL SKOR AKHIR": [93.5, 88.0, 71.5],
        "Rekomendasi Sistem": ["Pilih Utama", "Cadangan 1", "Perlu Evaluasi"]
    })
    st.dataframe(res_ranking, use_container_width=True)

# ---------------------------------------------------------
# MODUL 6: ANALISIS KOMPARASI HARGA PASAR & HET
# ---------------------------------------------------------
elif menu == "6. Analisis Komparasi Harga Pasar & HET":
    st.header("📈 Analisis Komparasi Harga Supplier vs HET vs Pasar")
    
    comp_df = pd.DataFrame({
        "Nama Komoditas": ["Beras Medium (Kg)", "Minyak Goreng (Ltr)", "Daging Ayam (Kg)", "Cabai Merah (Kg)", "Bawang Merah (Kg)"],
        "Harga Target Koperasi": [13000, 15000, 34000, 35000, 28000],
        "Penawaran Supplier A": [13500, 15200, 34500, 36000, 29000],
        "HET Bantul": [13100, 15000, 35000, 37000, 30000],
        "HET Sleman": [13200, 15000, 35000, 36500, 29500],
        "HET Gunungkidul": [13300, 15500, 36000, 38000, 31000],
        "HET Kulon Progo": [13100, 15000, 34800, 36000, 29000],
        "Hasil Survey Pasar": [13400, 15300, 35500, 37500, 30000]
    })
    
    st.dataframe(comp_df, use_container_width=True)
    
    st.subheader("📊 Grafik Perbandingan Harga Komoditas")
    st.line_chart(comp_df.set_index("Nama Komoditas")[["Harga Target Koperasi", "Penawaran Supplier A", "Hasil Survey Pasar"]])
