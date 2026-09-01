import streamlit as st
import pandas as pd

# Konfigurasi Halaman
st.set_page_config(page_title="Sistem Procurement Koperasi YK", layout="wide")

st.title("🛒 Sistem Evaluasi Supplier & Procurement")
st.subheader("Koperasi YK")
st.markdown("---")

# Tab Menu Aplikasi
tab1, tab2, tab3 = st.tabs(["📝 Input Evaluasi Supplier", "📊 Hasil & Peringkat", "ℹ️ Informasi Koperasi"])

with tab1:
    st.header("Form Evaluasi Supplier")
    
    with st.form("form_evaluasi"):
        nama_supplier = st.text_input("Nama Supplier / Vendor")
        kategori = st.selectbox("Kategori Barang", ["Bahan Pokok", "Atribut / Seragam", "ATK / Perlengkapan", "Lainnya"])
        
        st.subheader("Penilaian Kriteria (Skala 1 - 5)")
        kualitas = st.slider("1. Kualitas Barang/Produk", 1, 5, 3)
        harga = st.slider("2. Kesesuaian Harga", 1, 5, 3)
        pengiriman = st.slider("3. Ketepatan Waktu Pengiriman", 1, 5, 3)
        pelayanan = st.slider("4. Pelayanan & Respon Supplier", 1, 5, 3)
        
        catatan = st.text_area("Catatan / Evaluasi Tambahan")
        
        submitted = st.form_submit_button("Simpan Data Evaluasi")
        
        if submitted:
            total_skor = kualitas + harga + pengiriman + pelayanan
            rata_rata = total_skor / 4
            
            st.success(f"Data Supplier **{nama_supplier}** berhasil disimpan!")
            st.metric("Total Skor", f"{total_skor} / 20")
            st.metric("Rata-Rata Skor", f"{rata_rata:.2f} / 5.0")

with tab2:
    st.header("Ringkasan Evaluasi Supplier")
    st.info("Fitur simulasi data peringkat supplier.")
    
    # Data Sampel
    data_sample = {
        "Nama Supplier": ["CV Jaya Makmur", "PT Sinar Pangan", "UD Berkah Mandiri"],
        "Kategori": ["ATK / Perlengkapan", "Bahan Pokok", "Atribut / Seragam"],
        "Kualitas": [4, 5, 3],
        "Harga": [4, 4, 3],
        "Pengiriman": [5, 4, 2],
        "Pelayanan": [4, 5, 3],
        "Rata-Rata Skor": [4.25, 4.50, 2.75],
        "Status": ["Rekomendasi", "Sangat Direkomendasikan", "Perlu Evaluasi"]
    }
    
    df = pd.DataFrame(data_sample)
    st.dataframe(df, use_container_width=True)

with tab3:
    st.header("Informasi Koperasi YK")
    st.write("Aplikasi ini digunakan untuk mendukung sistem pengambilan keputusan pengadaan barang dan evaluasi supplier secara transparan dan terstruktur di Koperasi YK.")
