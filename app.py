import streamlit as st
import pandas as pd
import numpy as np
import io
import random
import string

# ---------------------------------------------------------
# CONFIG / SETUP HALAMAN
# ---------------------------------------------------------
st.set_page_config(
    page_title="Enterprise Procurement System - Koperasi YK",
    page_icon="🏭",
    layout="wide"
)

WEB_APP_URL = "https://sistem-procurement-koperasi-yk.streamlit.app"

# Helper Function: Generate Token Unik 32 Karakter
def buat_token(length=32):
    return ''.join(random.choices(string.ascii_letters + string.digits, k=length))

# ---------------------------------------------------------
# INITIALIZE SESSION STATE (DATA MASTER DEFAULT)
# ---------------------------------------------------------
if "df_supplier_state" not in st.session_state:
    st.session_state["df_supplier_state"] = pd.DataFrame(columns=[
        "KODE SUPPLIER", "NAMA SUPPLIER", "NO WA", 
        "JENIS BB 1", "JENIS BB 2", "JENIS BB 3", 
        "TOKEN", "LINK FORM", "PIC", "ALAMAT"
    ])

if "df_harga_bb" not in st.session_state:
    st.session_state["df_harga_bb"] = pd.DataFrame(columns=[
        "KODE SUPPLIER", "NAMA SUPPLIER", "NAMA BB", "HARGA PER SATUAN", "SATUAN"
    ])

if "df_dapur_state" not in st.session_state:
    st.session_state["df_dapur_state"] = pd.DataFrame(columns=[
        "KODE DAPUR", "NAMA DAPUR", "ALAMAT", "LATITUDE", "LONGITUDE"
    ])

# ---------------------------------------------------------
# SIDEBAR NAVIGATION
# ---------------------------------------------------------
st.sidebar.title("⚡ SPPG Engine")
st.sidebar.caption("Koperasi YK")
st.sidebar.markdown("---")

st.sidebar.subheader("🚀 Modul Operasional")
menu = st.sidebar.radio(
    "Pilih Modul:",
    [
        "📊 Dashboard & HET",
        "🏬 Kelola Data Dapur",
        "🤝 Data Supplier & Link Form",
        "💬 WA & PO Generator",
        "🚚 Matriks Jarak",
        "🎯 Scoring & Evaluasi",
        "📈 HET & Komparasi Pasar"
    ]
)

st.sidebar.markdown("---")
st.sidebar.caption("Status Data: 🟢 Sistem Evaluasi Supplier & Dapur.xlsb")


# ---------------------------------------------------------
# MODUL 1: DASHBOARD & HET
# ---------------------------------------------------------
if menu == "📊 Dashboard & HET":
    st.title("📊 Dashboard Utama & Pemantauan HET")
    st.write("Selamat datang di Enterprise Procurement System Koperasi YK.")
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Supplier", len(st.session_state["df_supplier_state"]))
    col2.metric("Total Dapur SPPG", len(st.session_state["df_dapur_state"]))
    col3.metric("Total Penawaran BB", len(st.session_state["df_harga_bb"]))


# ---------------------------------------------------------
# MODUL 2: KELOLA DATA DAPUR
# ---------------------------------------------------------
elif menu == "🏬 Kelola Data Dapur":
    st.title("🏬 Manajemen Data Dapur SPPG & Peta Sebaran")
    st.write("Kelola daftar dapur penerima manfaat dan lokasi titik koordinat.")
    
    st.dataframe(st.session_state["df_dapur_state"], use_container_width=True)


# ---------------------------------------------------------
# MODUL 3: DATA SUPPLIER & LINK FORM (UPLOAD EXCEL)
# ---------------------------------------------------------
elif menu == "🤝 Data Supplier & Link Form":
    st.subheader("🤝 Kelola Data Supplier & Link Web App")
    
    # --- FITUR UPLOAD FILE EXCEL / SPREADSHEET ---
    with st.expander("📂 **Upload Data Supplier dari File Excel / Google Sheets**", expanded=True):
        st.write("Unggah file `.xlsx`, `.xls`, atau `.csv` yang berisi daftar supplier kamu.")
        uploaded_file = st.file_uploader("Pilih file Excel Supplier:", type=["xlsx", "xls", "xlsb", "csv"])
        
        if uploaded_file is not None:
            try:
                # Membaca file
                if uploaded_file.name.endswith('.csv'):
                    df_upload = pd.read_csv(uploaded_file)
                else:
                    df_upload = pd.read_excel(uploaded_file)
                
                # Deteksi baris header jika terdapat judul/empty row di atasnya (seperti baris 3-4 di spreadsheet kamu)
                if "Supp Code" not in df_upload.columns and "Supplier Name" not in df_upload.columns:
                    for i in range(min(10, len(df_upload))):
                        row_vals = df_upload.iloc[i].astype(str).tolist()
                        if any("SUPP CODE" in str(v).upper() or "SUPPLIER NAME" in str(v).upper() for v in row_vals):
                            df_upload.columns = df_upload.iloc[i]
                            df_upload = df_upload.iloc[i+1:].reset_index(drop=True)
                            break

                # Cleaning nama kolom
                df_upload.columns = [str(c).strip() for c in df_upload.columns]
                
                # Mapping Nama Kolom dari Spreadsheet ke Format Sistem
                mapped_df = pd.DataFrame()
                
                # 1. Kode Supplier
                col_code = next((c for c in df_upload.columns if 'SUPP CODE' in c.upper() or 'KODE' in c.upper()), None)
                mapped_df["KODE SUPPLIER"] = df_upload[col_code].astype(str) if col_code else [f"SUP-{i+1:03d}" for i in range(len(df_upload))]
                
                # 2. Nama Supplier
                col_name = next((c for c in df_upload.columns if 'SUPPLIER NAME' in c.upper() or 'NAMA' in c.upper()), None)
                mapped_df["NAMA SUPPLIER"] = df_upload[col_name].astype(str) if col_name else "Tanpa Nama"
                
                # 3. Nomor WA / Phone (Auto-format ke format 628xxx)
                col_phone = next((c for c in df_upload.columns if 'PHONE' in c.upper() or 'WA' in c.upper() or 'TELP' in c.upper()), None)
                if col_phone:
                    def clean_phone(p):
                        p_str = str(p).replace("-", "").replace(" ", "").replace("+", "").split(".")[0].strip()
                        if p_str.startswith("0"):
                            return "62" + p_str[1:]
                        elif not p_str.startswith("62") and len(p_str) > 5:
                            return "62" + p_str
                        return p_str
                    mapped_df["NO WA"] = df_upload[col_phone].apply(clean_phone)
                else:
                    mapped_df["NO WA"] = "6285169796974"

                # 4. Kategori / Supply BB
                col_bb = next((c for c in df_upload.columns if 'SUPPLY BB' in c.upper() or 'JENIS' in c.upper() or 'KATEGORI' in c.upper()), None)
                mapped_df["JENIS BB 1"] = df_upload[col_bb].astype(str) if col_bb else "-"
                mapped_df["JENIS BB 2"] = "-"
                mapped_df["JENIS BB 3"] = "-"

                # 5. Generate Token Unik & Link Form Otomatis
                mapped_df["TOKEN"] = [buat_token() for _ in range(len(mapped_df))]
                mapped_df["LINK FORM"] = mapped_df["TOKEN"].apply(lambda t: f"{WEB_APP_URL}?token={t}")

                # Kolom Tambahan (PIC & Alamat)
                col_pic = next((c for c in df_upload.columns if 'PIC' in c.upper()), None)
                mapped_df["PIC"] = df_upload[col_pic].astype(str) if col_pic else "-"
                
                col_alamat = next((c for c in df_upload.columns if 'ALAMAT' in c.upper()), None)
                mapped_df["ALAMAT"] = df_upload[col_alamat].astype(str) if col_alamat else "-"

                # Menghapus baris yang kosong/nan
                mapped_df = mapped_df.dropna(subset=["NAMA SUPPLIER"]).reset_index(drop=True)

                st.success(f"✅ Berhasil membaca **{len(mapped_df)} data supplier** dari file!")
                st.dataframe(mapped_df[["KODE SUPPLIER", "NAMA SUPPLIER", "NO WA", "JENIS BB 1", "TOKEN"]].head(), use_container_width=True)

                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    if st.button("📥 Import & Timpa Data Supplier Baru", type="primary", use_container_width=True):
                        st.session_state["df_supplier_state"] = mapped_df
                        st.success("Data supplier berhasil di-import!")
                        st.rerun()
                with col_btn2:
                    if st.button("➕ Tambahkan ke Data Supplier Saat Ini", use_container_width=True):
                        st.session_state["df_supplier_state"] = pd.concat([st.session_state["df_supplier_state"], mapped_df], ignore_index=True)
                        st.success("Data supplier berhasil ditambahkan!")
                        st.rerun()

            except Exception as e:
                st.error(f"Gagal memproses file Excel: {e}")

    st.markdown("---")

    # --- TOMBOL EXPORT SELURUH DATA ---
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
        
    st.markdown("---")
    
    # --- TAMPILAN DAFTAR SUPPLIER ---
    df_sup = st.session_state["df_supplier_state"].copy()
    st.markdown(f"##### 📋 **Total Supplier Terdaftar: {len(df_sup)}**")
    
    if not df_sup.empty:
        for idx, row in df_sup.iterrows():
            tkn = row.get("TOKEN", "")
            form_url = f"{WEB_APP_URL}?token={tkn}" if pd.notna(tkn) and tkn != "" else ""
            
            with st.expander(f"🤝 **{row.get('KODE SUPPLIER', '')} - {row.get('NAMA SUPPLIER', '')}** (WA: +{row.get('NO WA', '-')})"):
                c1, c2 = st.columns([4, 1])
                with c1:
                    st.write(f"**Supply BB:** {row.get('JENIS BB 1', '-')} | **PIC:** {row.get('PIC', '-')}")
                    st.write(f"🔗 **Form Link:** [{form_url}]({form_url})")
                with c2:
                    if st.button("🗑️ Hapus", key=f"btn_del_sup_{idx}", type="secondary", use_container_width=True):
                        st.session_state["df_supplier_state"] = st.session_state["df_supplier_state"].drop(idx).reset_index(drop=True)
                        st.success("Supplier berhasil dihapus.")
                        st.rerun()
    else:
        st.info("Belum ada data supplier. Silakan upload file Excel di atas.")


# ---------------------------------------------------------
# MODUL 4: WA & PO GENERATOR
# ---------------------------------------------------------
elif menu == "💬 WA & PO Generator":
    st.title("💬 WhatsApp Broadcast & PO Generator")
    st.write("Kirim pesan blast penawaran harga ke supplier melalui WhatsApp.")


# ---------------------------------------------------------
# MODUL 5: MATRIKS JARAK
# ---------------------------------------------------------
elif menu == "🚚 Matriks Jarak":
    st.title("🚚 Hitung Matriks Jarak (Haversine/Maps)")
    st.write("Perhitungan jarak antar supplier dan dapur SPPG.")


# ---------------------------------------------------------
# MODUL 6: SCORING & EVALUASI
# ---------------------------------------------------------
elif menu == "🎯 Scoring & Evaluasi":
    st.title("🎯 Scoring & Evaluasi Supplier")
    st.write("Pemeringkatan supplier berdasarkan kriteria harga, jarak, dan kapasitas.")


# ---------------------------------------------------------
# MODUL 7: HET & KOMPARASI PASAR
# ---------------------------------------------------------
elif menu == "📈 HET & Komparasi Pasar":
    st.title("📈 Analisis HET & Komparasi Pasar")
    st.write("Analisis perbandingan harga penawaran supplier dengan Harga Eceran Tertinggi (HET).")


# ---------------------------------------------------------
# DEFAULT / FALLBACK
# ---------------------------------------------------------
else:
    st.info("Silakan pilih modul pada sidebar di sebelah kiri.")
