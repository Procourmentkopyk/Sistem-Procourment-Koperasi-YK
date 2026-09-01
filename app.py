import streamlit as st
import pandas as pd
import os

# Konfigurasi Halaman & Tema Modern
st.set_page_config(
    page_title="SPPG Procurement Engine | Koperasi YK",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS untuk Tampilan Modern & Canggih
st.markdown("""
<style>
    /* Styling Main Container */
    .main {
        background-color: #f8f9fa;
    }
    
    /* Custom Sidebar Header */
    .sidebar-title {
        font-size: 1.2rem;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 0.5rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    
    /* Modern Card Container */
    .nav-card {
        background: white;
        padding: 1.25rem;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
        border: 1px solid #e2e8f0;
        margin-bottom: 1rem;
        transition: all 0.2s ease-in-out;
    }
    .nav-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08);
    }
    
    /* Badge Status */
    .status-badge {
        display: inline-block;
        padding: 0.25em 0.6em;
        font-size: 75%;
        font-weight: 700;
        line-height: 1;
        text-align: center;
        white-space: nowrap;
        vertical-align: baseline;
        border-radius: 0.375rem;
        color: #065f46;
        background-color: #d1fae5;
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
                data_dict[sheet] = pd.read_excel(file_name, sheet_name=sheet)
            return data_dict, file_name
        except Exception as e:
            return None, str(e)
    return None, "File data tidak ditemukan."

data_excel, file_status = load_data()

# --- SIDEBAR COMPONENT ---
with st.sidebar:
    st.markdown('<div class="sidebar-title">⚡ <b>PROCUREMENT ENGINE</b></div>', unsafe_allow_html=True)
    st.caption("Koperasi YK — Sistem Evaluasi & Dapur SPPG")
    st.markdown("---")
    
    st.subheader("🎯 Navigasi Modul")
    
    menu_options = {
        "📊 Dashboard Utama": "1. Dashboard & HET",
        "🏬 Kelola Dapur SPPG": "2. Kelola Data Dapur SPPG",
        "💬 WA & PO Generator": "3. Penawaran WA & Update Harga PO",
        "🚛 Supplier & Matriks Jarak": "4. Data Supplier & Matriks Jarak",
        "🎯 Scoring & Evaluasi": "5. Penilaian & Penentuan Supplier per Dapur",
        "📈 Analisis HET & Pasar": "6. Analisis Komparasi Harga Pasar & HET"
    }
    
    selected_label = st.radio(
        "Pilih Modul Operasional:",
        list(menu_options.keys()),
        index=0
    )
    
    menu = menu_options[selected_label]
    
    st.markdown("---")
    st.markdown(f"**Status Koneksi Data:**")
    st.markdown(f'<span class="status-badge">🟢 {file_status}</span>', unsafe_allow_html=True)

# --- MAIN CONTENT HEADER ---
st.title("🏭 Sistem Evaluasi Supplier & Procurement Dapur SPPG")
st.markdown("##### Enterprise Procurement & Evaluation System — Koperasi YK")
st.markdown("---")

# ---------------------------------------------------------
# MODUL 1: DASHBOARD & HET
# ---------------------------------------------------------
if menu == "1. Dashboard & HET":
    st.subheader("📊 Executive Summary & Pemantauan HET")
    
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric(label="Total Dapur SPPG", value="12 Dapur", delta="Active")
    with m2:
        st.metric(label="Total Supplier Terdaftar", value="45 Supplier", delta="+3 Bulan ini")
    with m3:
        st.metric(label="Kategori Barang", value="6 Kategori", delta="Lengkap")
    with m4:
        st.metric(label="Rata-rata Ketepatan", value="94.2%", delta="+1.5% YoY")
        
    st.markdown("<br>", unsafe_allow_html=True)
    col_l, col_r = st.columns([2, 1])
    
    with col_l:
        st.markdown("### 🏆 Top 10 Supplier Terbaik (Performa Kombinasi)")
        data_top = {
            "Peringkat": [f"#{i}" for i in range(1, 11)],
            "Nama Supplier": [f"Supplier {chr(65+i)}" for i in range(10)],
            "Kategori Utama": ["Bahan Pokok", "Daging/Ayam", "Sayur", "Bahan Pokok", "Bumbu", "Sayur", "Daging/Ayam", "Bahan Pokok", "Buah", "Bumbu"],
            "Skor Evaluasi": [4.85, 4.78, 4.72, 4.65, 4.60, 4.58, 4.52, 4.49, 4.45, 4.40],
            "Rekomendasi": ["Sangat Direkomendasikan"] * 10
        }
        st.dataframe(pd.DataFrame(data_top), use_container_width=True, hide_index=True)
        
    with col_r:
        st.markdown("### 🌐 Portal HET Resmi")
        st.info("Akses langsung ke portal pemantauan harga pangan pemerintah:")
        st.markdown("- 🔗 [**Panel Harga Pangan BAPANAS**](https://panelharga.badandangan.go.id/)")
        st.markdown("- 🔗 [**Harga Bahan Pokok DIY (PIHPS)**](https://harga.jogjaprov.go.id/)")

# ---------------------------------------------------------
# MODUL 2: KELOLA DATA DAPUR SPPG
# ---------------------------------------------------------
elif menu == "2. Kelola Data Dapur SPPG":
    st.subheader("🏬 Manajemen Data Dapur SPPG")
    
    if data_excel and isinstance(data_excel, dict) and "Data Dapur" in data_excel:
        st.dataframe(data_excel["Data Dapur"], use_container_width=True)
    else:
        st.markdown("##### ➕ Form Pendaftaran Dapur SPPG Baru")
        with st.form("form_dapur_baru"):
            c1, c2 = st.columns(2)
            with c1:
                nama_dpr = st.text_input("Nama Dapur SPPG Baru")
                wil = st.selectbox("Wilayah Kab/Kota", ["Bantul", "Sleman", "Gunungkidul", "Kulon Progo", "Kota Yogyakarta"])
            with c2:
                kap = st.number_input("Kapasitas Produksi (Porsi)", min_value=100, step=100)
                pj = st.text_input("Penanggung Jawab Operasional")
            
            if st.form_submit_button("✨ Simpan & Tambahkan Dapur"):
                st.success(f"Dapur {nama_dpr} berhasil ditambahkan ke dalam database!")

# ---------------------------------------------------------
# MODUL 3: PENAWARAN WA & UPDATE HARGA PO
# ---------------------------------------------------------
elif menu == "3. Penawaran WA & Update Harga PO":
    st.subheader("📲 Otomasi Penawaran WhatsApp & Purchase Order (PO)")
    
    t_wa, t_po = st.tabs(["💬 Kirim WA Penawaran Mingguan", "📄 Generator PO Supplier"])
    
    with t_wa:
        c_w1, c_w2 = st.columns(2)
        with c_w1:
            sup_nama = st.text_input("Nama Supplier Target", "CV Jaya Makmur")
            sup_kat = st.selectbox("Kategori Penawaran", ["Bahan Pokok", "Sayuran & Buah", "Daging & Ayam", "Bumbu"])
            no_wa = st.text_input("Nomor WA Supplier (Format: 628xxx)", "628123456789")
        with c_w2:
            pesan = f"Halo {sup_nama},\nMohon update penawaran harga mingguan Koperasi YK untuk kategori *{sup_kat}* via link berikut: https://koperasi-yk.com/update-harga?sup={sup_nama.replace(' ', '%20')}\n\nTerima kasih."
            st.text_area("Pratinjau Pesan WA", pesan, height=120)
            link = f"https://wa.me/{no_wa}?text={pesan.replace(' ', '%20').replace('\n', '%0A')}"
            st.markdown(f"[🚀 **Kirim Pesan WhatsApp Sekarang**]({link})", unsafe_allow_html=True)

    with t_po:
        st.markdown("##### 📄 Preview Purchase Order (PO)")
        st.info("Fitur pembuatan PO otomatis berdasarkan hasil penilaian kriteria supplier terbaik.")
        st.button("🖨️ Cetak / Download PO (PDF)")

# ---------------------------------------------------------
# MODUL 4: DATA SUPPLIER & MATRIKS JARAK
# ---------------------------------------------------------
elif menu == "4. Data Supplier & Matriks Jarak":
    st.subheader("🚛 Database Supplier & Matriks Jarak Logistik")
    
    matriks_jarak = pd.DataFrame({
        "Nama Supplier": ["CV Jaya Makmur", "PT Sinar Pangan", "UD Berkah Mandiri"],
        "Kategori Komoditas": ["Bahan Pokok", "Daging/Ayam", "Sayur"],
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
    st.subheader("🎯 Penilaian & Evaluasi Multi-Kriteria Supplier")
    
    with st.expander("📝 Form Input Tanda Terima Dapur & Komplain Kualitas", expanded=True):
        with st.form("form_skor"):
            c1, c2, c3 = st.columns(3)
            with c1:
                st.selectbox("Dapur Target SPPG", ["Dapur SPPG Bantul Utama", "Dapur SPPG Sleman Timur"])
                st.selectbox("Supplier Evaluasi", ["CV Jaya Makmur", "PT Sinar Pangan", "UD Berkah Mandiri"])
            with c2:
                st.time_input("Jam Terima Barang", value=pd.to_datetime("06:15").time())
                st.selectbox("Status Kualitas / Komplain", ["Baik (Sesuai Spesifikasi)", "Komplain Minor", "Komplain Sedang", "Ditolak Total"])
            with c3:
                st.number_input("Harga Penawaran (Rp)", value=13500)
                st.number_input("Harga Target Koperasi (Rp)", value=13000)
            st.form_submit_button("⚡ Hitung Skor & Rekomendasi Opsi Terbaik")
        
    st.markdown("### 📋 Hasil Keputusan Supplier Terpilih")
    res = pd.DataFrame({
        "Rekomendasi": ["Peringkat 1 (Utama)", "Peringkat 2 (Cadangan)"],
        "Nama Supplier": ["CV Jaya Makmur", "PT Sinar Pangan"],
        "Skor Jarak (15%)": [95, 80],
        "Skor Harga (35%)": [90, 85],
        "Skor Kualitas (30%)": [100, 90],
        "Skor Waktu (20%)": [90, 100],
        "SKOR AKHIR": [93.5, 88.0],
        "Status Keputusan": ["PILIH UTAMA", "CADANGAN"]
    })
    st.dataframe(res, use_container_width=True, hide_index=True)

# ---------------------------------------------------------
# MODUL 6: ANALISIS KOMPARASI HARGA PASAR & HET
# ---------------------------------------------------------
elif menu == "6. Analisis Komparasi Harga Pasar & HET":
    st.subheader("📈 Analisis Komparasi Harga: Target vs Supplier vs HET vs Pasar")
    
    comp = pd.DataFrame({
        "Nama Komoditas": ["Beras Medium (Kg)", "Minyak Goreng (Ltr)", "Daging Ayam (Kg)", "Cabai Merah (Kg)"],
        "Target Koperasi": [13000, 15000, 34000, 35000],
        "Harga Supplier": [13500, 15200, 34500, 36000],
        "HET Bantul": [13100, 15000, 35000, 37000],
        "HET Sleman": [13200, 15000, 35000, 36500],
        "Survey Pasar": [13400, 15300, 35500, 37500]
    })
    
    st.dataframe(comp, use_container_width=True, hide_index=True)
    st.line_chart(comp.set_index("Nama Komoditas")[["Target Koperasi", "Harga Supplier", "Survey Pasar"]])
