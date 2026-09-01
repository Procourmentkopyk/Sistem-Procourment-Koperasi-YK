import io
import os
import random
import string
import urllib.parse
import numpy as np
import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# 1. KONFIGURASI HALAMAN & UTILITY
# ---------------------------------------------------------
st.set_page_config(
    page_title="SPPG Procurement Engine | Koperasi YK",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

WEB_APP_URL = "https://sistem-procurement-koperasi-yk.streamlit.app"


def buat_token(length=32):
    return "".join(
        random.choices(string.ascii_letters + string.digits, k=length)
    )


st.markdown(
    """
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
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------
# 2. FUNCTION LOAD DATA EXCEL / XLSB
# ---------------------------------------------------------
@st.cache_data
def load_data():
    file_name = "Sistem Evaluasi Supplier & Dapur.xlsb"
    if not os.path.exists(file_name):
        file_name = "Sistem Evaluasi Supplier & Dapur.xlsx"

    if os.path.exists(file_name):
        try:
            xls = pd.ExcelFile(file_name)
            data_dict = {sheet: pd.read_excel(file_name, sheet_name=sheet) for sheet in xls.sheet_names}
            return data_dict, file_name
        except Exception as e:
            return None, str(e)
    return None, "File data lokal tidak ditemukan (Menggunakan State Default)."


data_excel, file_status = load_data()

# ---------------------------------------------------------
# 3. INITIALIZE SESSION STATES
# ---------------------------------------------------------
# Master Item Bahan Baku (Pondasi Penawaran)
if "df_master_bb" not in st.session_state:
    st.session_state["df_master_bb"] = pd.DataFrame(
        [
            {"KODE": "KOP-BB-0001", "KATEGORI": "Beras", "ITEM BB": "Beras Premium", "SATUAN": "Kg"},
            {"KODE": "KOP-BB-0002", "KATEGORI": "Beras", "ITEM BB": "Beras Medium", "SATUAN": "Kg"},
            {"KODE": "KOP-BB-0006", "KATEGORI": "Telur Ayam", "ITEM BB": "Telur Ayam Ras Fresh", "SATUAN": "Kg"},
            {"KODE": "KOP-BB-0014", "KATEGORI": "Ayam", "ITEM BB": "Ayam Karkas (1kg isi 8)", "SATUAN": "Kg"},
            {"KODE": "KOP-BB-0020", "KATEGORI": "Minyak", "ITEM BB": "Minyak Goreng Sawit", "SATUAN": "Liter"},
        ]
    )

# State Dapur
if "df_dapur_state" not in st.session_state:
    if data_excel and isinstance(data_excel, dict) and "Data Dapur" in data_excel:
        raw_df = data_excel["Data Dapur"].copy()
        if "Unnamed" in str(raw_df.columns[0]):
            header_idx = raw_df[raw_df.apply(lambda row: row.astype(str).str.contains("NAMA DAPUR").any(), axis=1)].index
            if not header_idx.empty:
                idx = header_idx[0]
                raw_df.columns = raw_df.iloc[idx]
                raw_df = raw_df.iloc[idx + 1 :].reset_index(drop=True)
                raw_df = raw_df.dropna(how="all", subset=["NAMA DAPUR"])

        raw_df.columns = [str(c).strip().upper() for c in raw_df.columns]
        if "LONGTITUDE" in raw_df.columns:
            raw_df = raw_df.rename(columns={"LONGTITUDE": "LONGITUDE"})

        cols_to_keep = [c for c in ["KODE", "NAMA DAPUR", "ALAMAT", "PIC", "LATITUDE", "LONGITUDE", "KOTA/KABUPATEN"] if c in raw_df.columns]
        df_clean = raw_df[cols_to_keep].copy()
        df_clean["LATITUDE"] = pd.to_numeric(df_clean["LATITUDE"], errors="coerce")
        df_clean["LONGITUDE"] = pd.to_numeric(df_clean["LONGITUDE"], errors="coerce")
        st.session_state["df_dapur_state"] = df_clean.reset_index(drop=True)
    else:
        st.session_state["df_dapur_state"] = pd.DataFrame(
            {
                "KODE": ["PAKEM", "NGPLK", "SLMN4"],
                "NAMA DAPUR": ["PAKEM (Hargobinangun)", "NGEMPLAK (Umbulmartani 1)", "SLEMAN 4 (Triharjo)"],
                "ALAMAT": ["Jl. Kaliurang Km 22", "Jl. Kaliurang Km 15.5", "Jl. Letkol Subadri"],
                "PIC": ["SHINTA", "SHINTA", "CAHYO"],
                "LATITUDE": [-7.618382, -7.675789, -7.700475],
                "LONGITUDE": [110.426078, 110.417100, 110.342646],
                "KOTA/KABUPATEN": ["SLEMAN", "SLEMAN", "SLEMAN"],
            }
        )

# State Supplier (Dipastikan Kolom JENIS BB 1 s/d 6 Terinisialisasi)
if "df_supplier_state" not in st.session_state:
    cols_supplier = ["KODE SUPPLIER", "NAMA SUPPLIER", "NO WA", "TOKEN", "LINK FORM", "PIC", "ALAMAT"] + [f"JENIS BB {i}" for i in range(1, 7)]
    st.session_state["df_supplier_state"] = pd.DataFrame(
        [
            {
                "KODE SUPPLIER": "SUP-001",
                "NAMA SUPPLIER": "UD Tani Makmur",
                "NO WA": "6281234567890",
                "JENIS BB 1": "Ayam",
                "JENIS BB 2": "Telur Ayam",
                "JENIS BB 3": "-",
                "JENIS BB 4": "-",
                "JENIS BB 5": "-",
                "JENIS BB 6": "-",
                "TOKEN": "sample1",
                "LINK FORM": f"{WEB_APP_URL}?token=sample1",
                "PIC": "Budi",
                "ALAMAT": "Sleman",
            }
        ]
    )

# State Input Penawaran Harga
if "df_harga_bb" not in st.session_state:
    st.session_state["df_harga_bb"] = pd.DataFrame(
        columns=["KODE SUPPLIER", "NAMA SUPPLIER", "KODE_BB", "NAMA BB", "HARGA PER SATUAN", "SATUAN", "KATEGORI", "CATATAN"]
    )

# ---------------------------------------------------------
# 4. HANDLER FORM EXTERNAL VENDOR
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
        st.success(f"Selamat datang, **{supplier_info['NAMA SUPPLIER']}** (`{supplier_info['KODE SUPPLIER']}`)!")
    else:
        st.info("Form Penawaran Harga Barang Operasional Vendor External")

    df_master = st.session_state["df_master_bb"].copy()

    # Filter Master Barang jika parameter 'kat' di-pass di URL
    if kategori_diterima != "ALL":
        list_kat = [k.replace("_", " ").lower() for k in kategori_diterima.split(",")]
        df_master = df_master[df_master["KATEGORI"].str.lower().isin(list_kat)]

    with st.form("form_input_harga_supplier"):
        st.subheader("🛒 Isikan Harga Penawaran Sesuai Item Master")
        input_data = {}

        for idx, row in df_master.iterrows():
            c1, c2, c3 = st.columns([3, 1, 2])
            with c1:
                st.write(f"**{row['ITEM BB']}**\n*{row['KATEGORI']}*")
            with c2:
                st.write(f"Satuan: **{row['SATUAN']}**")
            with c3:
                input_data[row["KODE"]] = {
                    "harga": st.number_input("Harga (Rp)", min_value=0, step=500, value=0, key=f"hp_{row['KODE']}"),
                    "catatan": st.text_input("Catatan / Merek", key=f"ct_{row['KODE']}"),
                    "nama_bb": row["ITEM BB"],
                    "satuan": row["SATUAN"],
                    "kategori": row["KATEGORI"],
                }
            st.divider()

        submit_harga = st.form_submit_button("🚀 Kirim Seluruh Penawaran Harga", use_container_width=True)

        if submit_harga:
            kode_sup = supplier_info["KODE SUPPLIER"] if supplier_info is not None else "GUEST"
            nama_sup = supplier_info["NAMA SUPPLIER"] if supplier_info is not None else "Supplier External"

            new_entries = []
            for kode_bb, val in input_data.items():
                if val["harga"] > 0:
                    new_entries.append(
                        {
                            "KODE SUPPLIER": kode_sup,
                            "NAMA SUPPLIER": nama_sup,
                            "KODE_BB": kode_bb,
                            "NAMA BB": val["nama_bb"],
                            "HARGA PER SATUAN": val["harga"],
                            "SATUAN": val["satuan"],
                            "KATEGORI": val["kategori"],
                            "CATATAN": val["catatan"],
                        }
                    )

            if new_entries:
                df_new = pd.DataFrame(new_entries)
                # Hapus entry lama dari supplier yang sama agar data ter-update
                if not st.session_state["df_harga_bb"].empty:
                    st.session_state["df_harga_bb"] = st.session_state["df_harga_bb"][
                        st.session_state["df_harga_bb"]["KODE SUPPLIER"] != kode_sup
                    ]

                st.session_state["df_harga_bb"] = pd.concat([st.session_state["df_harga_bb"], df_new], ignore_index=True)
                st.balloons()
                st.success("Berhasil! Seluruh penawaran harga telah diperbarui.")
            else:
                st.warning("Mohon isi nominal harga minimal pada 1 barang.")

    st.stop()

# ---------------------------------------------------------
# 5. DIALOGS / POP-UP DAPUR
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
            new_row = pd.DataFrame(
                [{"KODE": kode, "NAMA DAPUR": nama, "ALAMAT": alamat, "PIC": pic, "LATITUDE": lat, "LONGITUDE": lon, "KOTA/KABUPATEN": kota}]
            )
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
# 6. SIDEBAR NAVIGASI
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
        "📈 HET & Komparasi Pasar",
    ]

    menu = st.radio("Pilih Modul:", modul_options, index=0)
    st.markdown("---")
    st.caption(f"Status Data: 🟢 {file_status}")

# ---------------------------------------------------------
# 7. HEADER UTAMA
# ---------------------------------------------------------
st.markdown(
    """
<div class="header-card">
    <div style="display: flex; justify-content: space-between; align-items: center;">
        <div>
            <h2 style="margin:0; color: white;">🏭 Enterprise Procurement System</h2>
            <p style="margin:0; opacity: 0.8;">Koperasi YK — Sistem Evaluasi Supplier & Dapur SPPG</p>
        </div>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------
# MODUL 1: DASHBOARD & HET
# ---------------------------------------------------------
if menu == "📊 Dashboard & HET":
    st.subheader("📊 Dashboard Utama & Rekapitulasi Harga Supplier")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Dapur SPPG", len(st.session_state["df_dapur_state"]), delta="Aktif")
    m2.metric("Total Supplier", len(st.session_state["df_supplier_state"]), delta="Terdaftar")
    m3.metric("Penawaran Masuk", len(st.session_state["df_harga_bb"]["KODE SUPPLIER"].unique()), delta="Vendor")
    m4.metric("Rata-rata Ketepatan", "94.2%", delta="+1.5%")

    st.markdown("---")
    st.markdown("##### 📋 **Matriks Komparasi Penawaran Vendor**")

    df_harga = st.session_state["df_harga_bb"]
    df_master = st.session_state["df_master_bb"]

    if not df_harga.empty:
        # Pivot Data Penawaran Harga
        pivot_df = df_harga.pivot(index="KODE_BB", columns="KODE SUPPLIER", values="HARGA PER SATUAN").reset_index()
        matrix_df = pd.merge(df_master, pivot_df, left_on="KODE", right_on="KODE_BB", how="left").drop(columns=["KODE_BB"], errors="ignore")
        matrix_df = matrix_df.fillna("-")
        st.dataframe(matrix_df, use_container_width=True)
    else:
        st.info("Belum ada penawaran harga yang masuk dari vendor.")
        st.dataframe(df_master, use_container_width=True)

# ---------------------------------------------------------
# MODUL 2: KELOLA DATA DAPUR
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
        st.map(map_data[["LATITUDE", "LONGITUDE"]].rename(columns={"LATITUDE": "lat", "LONGITUDE": "lon"}), zoom=10, use_container_width=True)
    else:
        st.warning("Data Koordinat belum terisi.")

    st.markdown("---")
    st.markdown("##### 📋 **Daftar Dapur SPPG Operasional**")

    if not df_dapur.empty:
        # Rendering menggunakan st.dataframe / Native Streamlit untuk performa & kestabilan UI
        st.dataframe(
            df_dapur,
            column_config={
                "LATITUDE": st.column_config.NumberColumn(format="%.5f"),
                "LONGITUDE": st.column_config.NumberColumn(format="%.5f"),
            },
            use_container_width=True,
        )

        st.caption("Gunakan aksi di bawah ini untuk memperbarui atau menghapus baris dapur:")
        c_sel, c_ed, c_del = st.columns([3, 1, 1])
        with c_sel:
            selected_dapur_idx = st.selectbox(
                "Pilih Dapur:",
                options=range(len(df_dapur)),
                format_func=lambda x: f"{df_dapur.iloc[x]['KODE']} - {df_dapur.iloc[x]['NAMA DAPUR']}",
            )
        with c_ed:
            st.write("")
            st.write("")
            if st.button("✏️ Edit Dapur", use_container_width=True):
                open_edit_dapur_dialog(selected_dapur_idx)
        with c_del:
            st.write("")
            st.write("")
            if st.button("🗑️ Hapus Dapur", type="primary", use_container_width=True):
                open_delete_dapur_dialog(selected_dapur_idx)
    else:
        st.info("Belum ada data dapur. Klik 'Tambah Dapur Baru' untuk menambahkan.")

# ---------------------------------------------------------
# MODUL 3: KELOLA DATA SUPPLIER & LINK WEB APP
# ---------------------------------------------------------
elif menu == "🤝 Data Supplier & Link Form":
    st.subheader("🤝 Kelola Data Supplier & Link Web App")

    with st.expander("📂 **Upload Data Supplier dari File Excel / Google Sheets**", expanded=True):
        uploaded_file = st.file_uploader("Pilih file Excel Supplier:", type=["xlsx", "xls", "xlsb", "csv"])

        if uploaded_file is not None:
            try:
                df_upload = pd.read_csv(uploaded_file) if uploaded_file.name.endswith(".csv") else pd.read_excel(uploaded_file)

                # Deteksi Header
                if "Supp Code" not in df_upload.columns and "Supplier Name" not in df_upload.columns:
                    for i in range(min(10, len(df_upload))):
                        row_vals = df_upload.iloc[i].astype(str).tolist()
                        if any("SUPP CODE" in str(v).upper() or "SUPPLIER NAME" in str(v).upper() for v in row_vals):
                            df_upload.columns = df_upload.iloc[i]
                            df_upload = df_upload.iloc[i + 1 :].reset_index(drop=True)
                            break

                df_upload.columns = [str(c).strip() for c in df_upload.columns]
                mapped_df = pd.DataFrame()

                col_code = next((c for c in df_upload.columns if "SUPP CODE" in c.upper() or "KODE" in c.upper()), None)
                mapped_df["KODE SUPPLIER"] = (
                    df_upload[col_code].astype(str) if col_code else [f"SUP-{i+1:03d}" for i in range(len(df_upload))]
                )

                col_name = next((c for c in df_upload.columns if "SUPPLIER NAME" in c.upper() or "NAMA" in c.upper()), None)
                mapped_df["NAMA SUPPLIER"] = df_upload[col_name].astype(str) if col_name else "Tanpa Nama"

                col_phone = next((c for c in df_upload.columns if any(k in c.upper() for k in ["PHONE", "WA", "TELP"])), None)
                if col_phone:

                    def clean_phone(p):
                        p_str = str(p).replace("-", "").replace(" ", "").replace("+", "").split(".")[0].strip()
                        return "62" + p_str[1:] if p_str.startswith("0") else (p_str if p_str.startswith("62") else "62" + p_str)

                    mapped_df["NO WA"] = df_upload[col_phone].apply(clean_phone)
                else:
                    mapped_df["NO WA"] = "6281234567890"

                # Masukkan Jenis BB (Satu per Satu ke Kolom 1..6)
                col_bb = next((c for c in df_upload.columns if any(k in c.upper() for k in ["SUPPLY BB", "JENIS", "KATEGORI"])), None)
                mapped_df["JENIS BB 1"] = df_upload[col_bb].astype(str) if col_bb else "-"
                for i in range(2, 7):
                    mapped_df[f"JENIS BB {i}"] = "-"

                mapped_df["TOKEN"] = [buat_token() for _ in range(len(mapped_df))]
                mapped_df["LINK FORM"] = mapped_df["TOKEN"].apply(lambda t: f"{WEB_APP_URL}?token={t}")

                col_pic = next((c for c in df_upload.columns if "PIC" in c.upper()), None)
                mapped_df["PIC"] = df_upload[col_pic].astype(str) if col_pic else "-"

                col_alamat = next((c for c in df_upload.columns if "ALAMAT" in c.upper()), None)
                mapped_df["ALAMAT"] = df_upload[col_alamat].astype(str) if col_alamat else "-"

                st.success(f"✅ Berhasil membaca **{len(mapped_df)} data supplier**!")
                st.dataframe(mapped_df[["KODE SUPPLIER", "NAMA SUPPLIER", "NO WA", "JENIS BB 1", "TOKEN"]].head(), use_container_width=True)

                c_imp1, c_imp2 = st.columns(2)
                with c_imp1:
                    if st.button("📥 Import & Timpa Data Supplier Baru", type="primary", use_container_width=True):
                        st.session_state["df_supplier_state"] = mapped_df
                        st.success("Data supplier di-import!")
                        st.rerun()
                with c_imp2:
                    if st.button("➕ Tambahkan ke Data Supplier Saat Ini", use_container_width=True):
                        st.session_state["df_supplier_state"] = pd.concat(
                            [st.session_state["df_supplier_state"], mapped_df], ignore_index=True
                        )
                        st.success("Data supplier ditambahkan!")
                        st.rerun()

            except Exception as e:
                st.error(f"Gagal memproses file Excel: {e}")

    # EXPORT EXCEL
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        st.session_state["df_harga_bb"].to_excel(writer, index=False, sheet_name="UPDATE HARGA BB")
        st.session_state["df_supplier_state"].to_excel(writer, index=False, sheet_name="Supplier Link")
        st.session_state["df_dapur_state"].to_excel(writer, index=False, sheet_name="Data Dapur")

    st.download_button(
        label="📥 Download Seluruh Master Data (Excel Multi-Sheet)",
        data=output.getvalue(),
        file_name="Master_Data_Procurement_System.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.markdown("---")
    df_sup = st.session_state["df_supplier_state"].copy()
    st.markdown(f"##### 📋 **Total Supplier Terdaftar: {len(df_sup)}**")

    if not df_sup.empty:
        for idx, row in df_sup.iterrows():
            tkn = row.get("TOKEN", "")
            form_url = f"{WEB_APP_URL}?token={tkn}" if pd.notna(tkn) and tkn != "" else ""

            with st.expander(f"🤝 **{row.get('KODE SUPPLIER', '')} - {row.get('NAMA SUPPLIER', '')}** (WA: +{row.get('NO WA', '-')})"):
                c1, c2 = st.columns([4, 1])
                with c1:
                    st.write(f"**Supply BB Utama:** {row.get('JENIS BB 1', '-')} | **PIC:** {row.get('PIC', '-')}")
                    st.write(f"🔗 **Form Link:** [{form_url}]({form_url})")
                with c2:
                    if st.button("🗑️ Hapus", key=f"btn_del_sup_{idx}", type="secondary", use_container_width=True):
                        st.session_state["df_supplier_state"] = st.session_state["df_supplier_state"].drop(idx).reset_index(drop=True)
                        st.success("Supplier dihapus.")
                        st.rerun()
    else:
        st.info("Belum ada data supplier.")

# ---------------------------------------------------------
# MODUL 4: WA & PO GENERATOR
# ---------------------------------------------------------
elif menu == "💬 WA & PO Generator":
    st.subheader("💬 WA Message & PO Link Generator")
    df_sup = st.session_state["df_supplier_state"]

    if df_sup.empty:
        st.warning("⚠️ Data supplier kosong.")
    else:
        LIST_KATEGORI = [
            "Ayam",
            "Beras",
            "Buah",
            "Cookies",
            "Daging",
            "Ikan",
            "Keju",
            "Minyak",
            "Olahan",
            "Sayur",
            "Sembako",
            "Susu",
            "Tahu",
            "Tempe",
            "Telur Ayam",
            "Telur Bebek",
            "Telur Puyuh",
        ]

        st.markdown("##### 1️⃣ **Pilih Supplier & Atur Kategori Barang**")
        list_supplier_display = [f"{r['KODE SUPPLIER']} - {r['NAMA SUPPLIER']}" for _, r in df_sup.iterrows()]
        selected_sup_str = st.selectbox("Pilih Supplier Target:", list_supplier_display)

        idx_sup = list_supplier_display.index(selected_sup_str)
        row_sup = df_sup.iloc[idx_sup]

        kat_existing = []
        for col_bb in [f"JENIS BB {i}" for i in range(1, 7)]:
            val = row_sup.get(col_bb, "-")
            if pd.notna(val) and str(val).strip() not in ["-", "", "nan", "None"]:
                matched = next((k for k in LIST_KATEGORI if k.upper() == str(val).strip().upper()), None)
                if matched:
                    kat_existing.append(matched)

        selected_kategori = st.multiselect(
            "Pilih Kategori Barang Ditawarkan (Maksimal 6):",
            options=LIST_KATEGORI,
            default=kat_existing,
            max_selections=6,
            key=f"ms_kat_{idx_sup}",
        )

        if st.button("💾 Simpan Kategori Supplier", type="primary"):
            for i in range(1, 7):
                col_name = f"JENIS BB {i}"
                val_to_save = selected_kategori[i - 1] if i <= len(selected_kategori) else "-"
                st.session_state["df_supplier_state"].at[idx_sup, col_name] = val_to_save
            st.success("Kategori disimpan!")
            st.rerun()

        st.markdown("---")
        st.markdown("##### 2️⃣ **Preview & Edit Draft WhatsApp**")

        token = row_sup.get("TOKEN", "")
        kat_param = ",".join([k.replace(" ", "_") for k in selected_kategori]) if selected_kategori else "ALL"
        link_form_khusus = f"{WEB_APP_URL}?token={token}&kat={kat_param}"

        no_wa = str(row_sup.get("NO WA", "")).replace("+", "").replace(" ", "").replace("-", "").split(".")[0]
        nama_supplier = row_sup.get("NAMA SUPPLIER", "Bapak/Ibu Vendor")
        kategori_txt = ", ".join(selected_kategori) if selected_kategori else "Bahan Baku"

        default_pesan = (
            f"Halo *{nama_supplier}*,\n\n"
            f"Kami dari tim Procurement Koperasi YK/SPPG. "
            f"Mohon untuk memperbarui update harga penawaran harian/mingguan kategori *[{kategori_txt}]* "
            f"melalui link form resmi berikut ini:\n\n"
            f"🔗 {link_form_khusus}\n\n"
            f"Terima kasih atas kerja samanya."
        )

        pesan_edited = st.text_area("Draft Pesan WhatsApp:", value=default_pesan, height=180, key=f"txt_wa_{idx_sup}")

        encoded_msg = urllib.parse.quote(pesan_edited)
        wa_url = f"https://wa.me/{no_wa}?text={encoded_msg}"

        st.markdown(
            f'<a href="{wa_url}" target="_blank" style="text-decoration:none;">'
            f'<div style="background-color:#25D366; color:white; padding:12px; border-radius:8px; text-align:center; font-weight:bold;">'
            f"📲 Kirim via WhatsApp (+{no_wa})</div></a>",
            unsafe_allow_html=True,
        )

# ---------------------------------------------------------
# MODUL OPERASIONAL LAINNYA (PLACEHOLDER)
# ---------------------------------------------------------
elif menu == "🚛 Matriks Jarak":
    st.subheader("🚛 Matriks Jarak & Rute Distribusi Dapur SPPG")
elif menu == "🎯 Scoring & Evaluasi":
    st.subheader("🎯 Scoring & Evaluasi Kinerja Supplier")
elif menu == "📈 HET & Komparasi Pasar":
    st.subheader("📈 Pemantauan Harga HET & Komparasi Pasar")
