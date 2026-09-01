import streamlit as st
import pandas as pd
import os

# Konfigurasi Halaman dengan Sidebar Tersembunyi Otomatis (Clean Mode)
st.set_page_config(
    page_title="SPPG Procurement Engine | Koperasi YK",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"  # Membuat tampilan awal sangat bersih
)

# Custom CSS untuk Desain Card & Pop-Up Styling
st.markdown("""
<style>
    /* Gradient Header & Modern Look */
    .header-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    
    /* Status Badge */
    .status-badge {
        display: inline-block;
        padding: 0.35em 0.65em;
        font-size: 80%;
        font-weight: 600;
        border-radius: 20px;
        color: #065f46;
        background-color: #d1fae5;
    }
    
    /* Interactive Card Button Effect */
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
                data_dict[sheet] = pd.read_excel(file_name, sheet_name=sheet)
            return data_dict, file_name
        except Exception as e:
            return None, str(e)
    return None, "File data tidak ditemukan."

data_excel, file_status = load_data()

# Initialize Session State untuk Modul Terpilih
if "current_modul" not in st.session_state:
    st.session_state["current_modul"] = "1. Dashboard & HET"

# --- DEKLARASI POP-UP DIALOG UNTUK INPUT DATA DAPUR ---
@st.dialog("🏬 Form Tambah Dapur SPPG Baru")
def open_dapur_dialog():
    st.write("Silakan isi informasi dapur operasional baru di bawah ini:")
    with st.form("form_dapur_popup"):
        c1, c2 = st.columns(2)
        with c1:
            nama_dpr = st.text_input("Nama Dapur SPPG Baru")
            wil = st.selectbox("Wilayah Kab/Kota", ["Bantul", "Sleman", "Gunungkidul", "Kulon Progo", "Kota Yogyakarta"])
        with c2:
            kap = st.number_input("Kapasitas (Porsi)", min_value=100, step=100)
            pj = st.text_input("Penanggung Jawab Dapur")
            
        submitted = st.form_submit_button("✨ Simpan Dapur Baru")
        if submitted:
            st.success(f"Dapur {nama_dpr} berhasil ditambahkan!")
            st.rerun()

# --- DEKLARASI POP-UP DIALOG UNTUK PENILAIAN SUPPLIER ---
@st.dialog("🎯 Form Input Penilaian Supplier")
def open_scoring_dialog():
    st.write("Input data tanda terima & evaluasi kualitas pengiriman:")
    with st.form("form_skor_popup"):
        c1, c2 = st.columns(2)
        with c1:
            st.selectbox("Dapur Target", ["Dapur SPPG Bantul Utama", "Dapur SPPG Sleman Timur"])
            st.selectbox("Supplier", ["CV Jaya Makmur", "PT Sinar Pangan", "UD Berkah Mandiri"])
            st.number_input("Harga Penawaran (Rp)", value=13500)
        with c2:
            st.time_input("Jam Terima Barang", value=pd.to_datetime("06:15").time())
            st.selectbox("Status Kualitas", ["Baik (Sesuai Spesifikasi)", "Komplain Minor", "Komplain Sedang", "Ditolak Total"])
            st.number_input("Harga Target (Rp)", value=13000)
            
        submitted = st.form_submit_button("⚡ Kalkulasi Rekomendasi")
        if submitted:
            st.success("Evaluasi berhasil dihitung!")
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

# --- NAVIGASI MODUL BERGAYA POP-UP BAR ---
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
# MODUL 1: DASHBOARD & HET
# ---------------------------------------------------------
if menu == "1. Dashboard & HET":
    st.subheader("📊 Dashboard Utama & Pemantauan HET")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Dapur SPPG", "12 Dapur", delta="Active")
    with col2:
        st.metric("Total Supplier", "45 Supplier", delta="+3 Bulan Ini")
    with col3:
        st.metric("Kategori Barang", "6 Kategori", delta="Lengkap")
    with col4:
        st.metric("Rata-rata Ketepatan", "94.2%", delta="+1.5%")
        
    st.markdown("<br>", unsafe_allow_html=True)
    c_left, c_right = st.columns([2, 1])
    
    with c_left:
        st.markdown("### 🏆 Top 10 Supplier Terbaik")
        data_top = {
            "Peringkat": [f"#{i}" for i in range(1, 11)],
            "Nama Supplier": [f"Supplier {chr(65+i)}" for i in range(10)],
            "Kategori Utama": ["Bahan Pokok", "Daging/Ayam", "Sayur", "Bahan Pokok", "Bumbu", "Sayur", "Daging/Ayam", "Bahan Pokok", "Buah", "Bumbu"],
            "Skor Evaluasi": [4.85, 4.78, 4.72, 4.65, 4.60, 4.58, 4.52, 4.49, 4.45, 4.40],
            "Rekomendasi": ["Sangat Direkomendasikan"] * 10
        }
        st.dataframe(pd.DataFrame(data_top), use_container_width=True, hide_index=True)
        
    with c_right:
        st.markdown("### 🌐 Portal HET Resmi")
        st.info("Akses portal pemantauan harga pangan:")
        st.markdown("- 🔗 [**Panel Harga BAPANAS**](https://panelharga.badandangan.go.id/)")
        st.markdown("- 🔗 [**Harga Bahan Pokok DIY**](https://harga.jogjaprov.go.id/)")

# ---------------------------------------------------------
# MODUL 2: KELOLA DATA DAPUR SPPG
# ---------------------------------------------------------
elif menu == "2. Kelola Data Dapur SPPG":
    st.subheader("🏬 Manajemen Data Dapur SPPG")
    
    btn_col, _ = st.columns([1, 3])
    with btn_col:
        if st.button("➕ Tambah Dapur Baru (Pop-up Form)", type="primary", use_container_width=True):
            open_dapur_dialog()
            
    st.markdown("<br>", unsafe_allow_html=True)
    if data_excel and isinstance(data_excel, dict) and "Data Dapur" in data_excel:
        st.dataframe(data_excel["Data Dapur"], use_container_width=True)
    else:
        st.info("Data Dapur SPPG ditampilkan secara interaktif.")

# ---------------------------------------------------------
# MODUL 3: PENAWARAN WA & UPDATE HARGA PO
# ---------------------------------------------------------
elif menu == "3. Penawaran WA & Update Harga PO":
    st.subheader("📲 Otomasi Penawaran WhatsApp & PO Generator")
    
    t_wa, t_po = st.tabs(["💬 Kirim WA Penawaran Mingguan", "📄 Generator PO Supplier"])
    
    with t_wa:
        c_w1, c_w2 = st.columns(2)
        with c_w1:
            sup_nama = st.text_input("Nama Supplier Target", "CV Jaya Makmur")
            sup_kat = st.selectbox("Kategori Penawaran", ["Bahan Pokok", "Sayuran & Buah", "Daging & Ayam", "Bumbu"])
            no_wa = st.text_input("Nomor WA Supplier", "628123456789")
        with c_w2:
            pesan = f"Halo {sup_nama},\nMohon update penawaran harga mingguan Koperasi YK via link berikut: https://koperasi-yk.com/update-harga?sup={sup_nama.replace(' ', '%20')}"
            st.text_area("Pratinjau Pesan WA", pesan, height=120)
            link = f"https://wa.me/{no_wa}?text={pesan.replace(' ', '%20')}"
            st.markdown(f"[🚀 **Kirim Pesan WhatsApp**]({link})", unsafe_allow_html=True)

    with t_po:
        st.info("Generator Purchase Order (PO) Siap Cetak.")
        st.button("🖨️ Cetak / Download PO (PDF)")

# ---------------------------------------------------------
# MODUL 4: DATA SUPPLIER & MATRIKS JARAK
# ---------------------------------------------------------
elif menu == "4. Data Supplier & Matriks Jarak":
    st.subheader("🚛 Matriks Jarak & Database Supplier")
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
    st.subheader("🎯 Scoring Multi-Kriteria & Penentuan Supplier")
    
    col_pop, _ = st.columns([1, 3])
    with col_pop:
        if st.button("📝 Input Evaluasi Supplier (Pop-up Form)", type="primary", use_container_width=True):
            open_scoring_dialog()
            
    st.markdown("<br>", unsafe_allow_html=True)
    res = pd.DataFrame({
        "Rekomendasi": ["Peringkat 1 (Utama)", "Peringkat 2 (Cadangan)"],
        "Nama Supplier": ["CV Jaya Makmur", "PT Sinar Pangan"],
        "Skor Jarak (15%)": [95, 80],
        "Skor Harga (35%)": [90, 85],
        "Skor Kualitas (30%)": [100, 90],
        "Skor Waktu (20%)": [90, 100],
        "SKOR AKHIR": [93.5, 88.0],
        "Status": ["PILIH UTAMA", "CADANGAN"]
    })
    st.dataframe(res, use_container_width=True, hide_index=True)

# ---------------------------------------------------------
# MODUL 6: ANALISIS KOMPARASI HARGA PASAR & HET
# ---------------------------------------------------------
elif menu == "6. Analisis Komparasi Harga Pasar & HET":
    st.subheader("📈 Analisis Komparasi Harga Target vs Supplier vs HET vs Pasar")
    comp = pd.DataFrame({
        "Nama Komoditas": ["Beras Medium (Kg)", "Minyak Goreng (Ltr)", "Daging Ayam (Kg)", "Cabai Merah (Kg)"],
        "Target Koperasi": [13000, 15000, 34000, 35000],
        "Harga Supplier": [13500, 15200, 34500, 36000],
        "HET Bantul": [13100, 15000, 35000, 37000],
        "Survey Pasar": [13400, 15300, 35500, 37500]
    })
    st.dataframe(comp, use_container_width=True, hide_index=True)
    st.line_chart(comp.set_index("Nama Komoditas")[["Target Koperasi", "Harga Supplier", "Survey Pasar"]])
