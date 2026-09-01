import io
import os
import random
import string
import urllib.parse
import folium
import pandas as pd
import streamlit as st
from streamlit_folium import st_folium

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
# 2. INITIALIZE SESSION STATES
# ---------------------------------------------------------
# State Master Bahan Baku (BB)
if "df_master_bb" not in st.session_state:
    st.session_state["df_master_bb"] = pd.DataFrame(
        [
            {
                "KODE": "KOP-BB-0001",
                "KATEGORI": "Beras",
                "ITEM BB": "Beras Premium",
                "SATUAN": "Kg",
                "HET": 14900,
            },
            {
                "KODE": "KOP-BB-0002",
                "KATEGORI": "Beras",
                "ITEM BB": "Beras Medium",
                "SATUAN": "Kg",
                "HET": 12500,
            },
            {
                "KODE": "KOP-BB-0006",
                "KATEGORI": "Telur Ayam",
                "ITEM BB": "Telur Ayam Ras Fresh",
                "SATUAN": "Kg",
                "HET": 28000,
            },
            {
                "KODE": "KOP-BB-0014",
                "KATEGORI": "Ayam",
                "ITEM BB": "Ayam Karkas (1kg isi 8)",
                "SATUAN": "Kg",
                "HET": 36000,
            },
            {
                "KODE": "KOP-BB-0020",
                "KATEGORI": "Minyak",
                "ITEM BB": "Minyak Goreng Sawit",
                "SATUAN": "Liter",
                "HET": 15700,
            },
            {
                "KODE": "KOP-BB-0025",
                "KATEGORI": "Daging",
                "ITEM BB": "Daging Sapi Murni",
                "SATUAN": "Kg",
                "HET": 130000,
            },
            {
                "KODE": "KOP-BB-0030",
                "KATEGORI": "Sayur",
                "ITEM BB": "Wortel Segar",
                "SATUAN": "Kg",
                "HET": 12000,
            },
        ]
    )

# State Data Dapur SPPG
if "df_dapur_state" not in st.session_state:
    st.session_state["df_dapur_state"] = pd.DataFrame(
        [
            {
                "KODE": "PAKEM",
                "NAMA DAPUR": "PAKEM (Hargobinangun)",
                "ALAMAT": "Jl. Kaliurang Km 22, Hargobinangun, Pakem, Sleman",
                "PIC": "SHINTA",
                "NO TELP": "6281234567891",
                "LATITUDE": -7.618382,
                "LONGITUDE": 110.426078,
                "KOTA/KABUPATEN": "SLEMAN",
            },
            {
                "KODE": "NGPLK",
                "NAMA DAPUR": "NGEMPLAK (Umbulmartani 1)",
                "ALAMAT": "Jl. Kaliurang Km 15.5, Umbulmartani, Ngemplak, Sleman",
                "PIC": "SHINTA",
                "NO TELP": "6281234567892",
                "LATITUDE": -7.675789,
                "LONGITUDE": 110.417100,
                "KOTA/KABUPATEN": "SLEMAN",
            },
            {
                "KODE": "SLMN4",
                "NAMA DAPUR": "SLEMAN 4 (Triharjo)",
                "ALAMAT": "Jl. Letkol Subadri, Triharjo, Sleman",
                "PIC": "CAHYO",
                "NO TELP": "6281234567893",
                "LATITUDE": -7.700475,
                "LONGITUDE": 110.342646,
                "KOTA/KABUPATEN": "SLEMAN",
            },
        ]
    )

# State Supplier
if "df_supplier_state" not in st.session_state:
    cols_supplier = [
        "KODE SUPPLIER",
        "NAMA SUPPLIER",
        "NO WA",
        "TOKEN",
        "LINK FORM",
        "PIC",
        "ALAMAT",
    ] + [f"JENIS BB {i}" for i in range(1, 7)]
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
        columns=[
            "KODE SUPPLIER",
            "NAMA SUPPLIER",
            "KODE_BB",
            "NAMA BB",
            "HARGA PER SATUAN",
            "SATUAN",
            "KATEGORI",
            "CATATAN",
        ]
    )

# ---------------------------------------------------------
# 3. HANDLER FORM EXTERNAL VENDOR
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
        st.success(
            f"Selamat datang, **{supplier_info['NAMA SUPPLIER']}**"
            f" (`{supplier_info['KODE SUPPLIER']}`)!"
        )
    else:
        st.info("Form Penawaran Harga Barang Operasional Vendor External")

    df_master = st.session_state["df_master_bb"].copy()

    # Filter Kategori Aman (Fallback)
    if kategori_diterima != "ALL" and str(kategori_diterima).strip() != "":
        list_kat = [
            k.replace("_", " ").strip().lower()
            for k in kategori_diterima.split(",")
        ]
        df_filtered = df_master[
            df_master["KATEGORI"].str.lower().isin(list_kat)
        ]
        if not df_filtered.empty:
            df_master = df_filtered

    with st.form("form_input_harga_supplier"):
        st.subheader("🛒 Isikan Harga Penawaran Sesuai Item Master")
        st.caption(f"Menampilkan {len(df_master)} item barang penawaran.")

        input_data = {}

        for idx, row in df_master.iterrows():
            c1, c2, c3 = st.columns([3, 1, 2])
            with c1:
                st.write(f"**{row['ITEM BB']}**")
                st.caption(
                    f"Kategori: {row['KATEGORI']} | Kode: {row['KODE']}"
                )
            with c2:
                st.write(f"Satuan: **{row['SATUAN']}**")
            with c3:
                input_data[row["KODE"]] = {
                    "harga": st.number_input(
                        "Harga Penawaran (Rp)",
                        min_value=0,
                        step=500,
                        value=0,
                        key=f"hp_{row['KODE']}",
                    ),
                    "catatan": st.text_input(
                        "Catatan / Merek", key=f"ct_{row['KODE']}"
                    ),
                    "nama_bb": row["ITEM BB"],
                    "satuan": row["SATUAN"],
                    "kategori": row["KATEGORI"],
                }
            st.divider()

        submit_harga = st.form_submit_button(
            "🚀 Kirim Seluruh Penawaran Harga", use_container_width=True
        )

        if submit_harga:
            kode_sup = (
                supplier_info["KODE SUPPLIER"]
                if supplier_info is not None
                else "GUEST"
            )
            nama_sup = (
                supplier_info["NAMA SUPPLIER"]
                if supplier_info is not None
                else "Supplier External"
            )

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
                if not st.session_state["df_harga_bb"].empty:
                    st.session_state["df_harga_bb"] = st.session_state[
                        "df_harga_bb"
                    ][
                        st.session_state["df_harga_bb"]["KODE SUPPLIER"]
                        != kode_sup
                    ]

                st.session_state["df_harga_bb"] = pd.concat(
                    [st.session_state["df_harga_bb"], df_new],
                    ignore_index=True,
                )
                st.balloons()
                st.success(
                    "Berhasil! Seluruh penawaran harga telah tersimpan."
                )
            else:
                st.warning("Mohon isi nominal harga minimal pada 1 barang.")

    st.stop()

# ---------------------------------------------------------
# 4. SIDEBAR NAVIGASI
# ---------------------------------------------------------
with st.sidebar:
    st.markdown("### ⚡ **SPPG Engine**")
    st.caption("Koperasi YK")
    st.markdown("---")
    st.subheader("🚀 Modul Operasional")

    modul_options = [
        "📊 Dashboard & HET",
        "📦 Master Bahan Baku (BB)",
        "🏬 Kelola Data Dapur",  # Modul Dapur Lengkap CRUD & Peta
        "🤝 Data Supplier & Link Form",
        "💬 WA & PO Generator",
        "🎯 Scoring & Evaluasi",
    ]

    menu = st.radio("Pilih Modul:", modul_options, index=2)
    st.markdown("---")

# ---------------------------------------------------------
# 5. HEADER UTAMA
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
# MODUL DAPUR LENGKAP (CRUD + PETA FOLIUM)
# ---------------------------------------------------------
if menu == "🏬 Kelola Data Dapur":
    st.subheader("🏬 Manajemen Data Dapur SPPG Operasional")

    tab_view, tab_add, tab_edit_delete = st.tabs(
        ["🗺️ Peta & Daftar Dapur", "➕ Tambah Dapur Baru", "⚙️ Kelola (Edit/Hapus)"]
    )

    # TAB 1: PETA & DAFTAR DAPUR
    with tab_view:
        col_list, col_peta = st.columns([1, 1])

        with col_list:
            st.markdown("##### 📋 **Daftar Dapur Terdaftar**")
            st.dataframe(
                st.session_state["df_dapur_state"][
                    ["KODE", "NAMA DAPUR", "PIC", "KOTA/KABUPATEN"]
                ],
                use_container_width=True,
            )

        with col_peta:
            st.markdown("##### 📍 **Peta Persebaran Dapur SPPG**")
            df_dapur = st.session_state["df_dapur_state"]

            if not df_dapur.empty:
                avg_lat = df_dapur["LATITUDE"].mean()
                avg_lon = df_dapur["LONGITUDE"].mean()
                m = folium.Map(location=[avg_lat, avg_lon], zoom_start=11)

                for idx, row in df_dapur.iterrows():
                    popup_html = f"""
                    <b>{row['NAMA DAPUR']}</b><br>
                    PIC: {row['PIC']}<br>
                    Alamat: {row['ALAMAT']}
                    """
                    folium.Marker(
                        location=[row["LATITUDE"], row["LONGITUDE"]],
                        popup=popup_html,
                        tooltip=row["NAMA DAPUR"],
                        icon=folium.Icon(color="red", icon="cutlery"),
                    ).add_to(m)

                st_folium(m, width=500, height=350)

    # TAB 2: TAMBAH DAPUR
    with tab_add:
        with st.form("form_tambah_dapur"):
            st.write("➕ **Form Tambah Dapur Operasional Baru**")
            c1, c2 = st.columns(2)
            with c1:
                kode_dapur = st.text_input("Kode Dapur (misal: YOGYA1)")
                nama_dapur = st.text_input("Nama Dapur")
                pic_dapur = st.text_input("Nama PIC Dapur")
                telp_dapur = st.text_input("No WhatsApp/Telp PIC")
            with c2:
                alamat_dapur = st.text_area("Alamat Lengkap")
                kota_dapur = st.text_input(
                    "Kota/Kabupaten", value="SLEMAN"
                )
                lat_dapur = st.number_input(
                    "Latitude", value=-7.700000, format="%.6f"
                )
                lon_dapur = st.number_input(
                    "Longitude", value=110.350000, format="%.6f"
                )

            if st.form_submit_button("✨ Simpan Data Dapur Baru"):
                if kode_dapur and nama_dapur:
                    new_dapur = pd.DataFrame(
                        [
                            {
                                "KODE": kode_dapur,
                                "NAMA DAPUR": nama_dapur,
                                "ALAMAT": alamat_dapur,
                                "PIC": pic_dapur,
                                "NO TELP": telp_dapur,
                                "LATITUDE": lat_dapur,
                                "LONGITUDE": lon_dapur,
                                "KOTA/KABUPATEN": kota_dapur,
                            }
                        ]
                    )
                    st.session_state["df_dapur_state"] = pd.concat(
                        [st.session_state["df_dapur_state"], new_dapur],
                        ignore_ignore=True,
                    )
                    st.success("Data Dapur Baru Berhasil Ditambahkan!")
                    st.rerun()
                else:
                    st.warning("Kode Dapur dan Nama Dapur wajib diisi.")

    # TAB 3: KELOLA (EDIT & HAPUS DAPUR)
    with tab_edit_delete:
        st.write("⚙️ **Edit atau Hapus Data Dapur**")
        df_d = st.session_state["df_dapur_state"]

        if not df_d.empty:
            dapur_pilihan = st.selectbox(
                "Pilih Dapur yang akan Diubah/Dihapus:",
                df_d["NAMA DAPUR"].tolist(),
            )
            data_selected = df_d[df_d["NAMA DAPUR"] == dapur_pilihan].iloc[0]

            with st.form("form_edit_dapur"):
                c1, c2 = st.columns(2)
                with c1:
                    e_kode = st.text_input(
                        "Kode Dapur", value=data_selected["KODE"]
                    )
                    e_nama = st.text_input(
                        "Nama Dapur", value=data_selected["NAMA DAPUR"]
                    )
                    e_pic = st.text_input(
                        "PIC", value=data_selected["PIC"]
                    )
                with c2:
                    e_alamat = st.text_area(
                        "Alamat", value=data_selected["ALAMAT"]
                    )
                    e_lat = st.number_input(
                        "Latitude",
                        value=float(data_selected["LATITUDE"]),
                        format="%.6f",
                    )
                    e_lon = st.number_input(
                        "Longitude",
                        value=float(data_selected["LONGITUDE"]),
                        format="%.6f",
                    )

                c_edit, c_del = st.columns([1, 1])
                with c_edit:
                    btn_update = st.form_submit_button("✏️ Simpan Perubahan")
                with c_del:
                    btn_delete = st.form_submit_button("🗑️ Hapus Dapur Ini")

                if btn_update:
                    idx = df_d[df_d["NAMA DAPUR"] == dapur_pilihan].index[0]
                    st.session_state["df_dapur_state"].at[idx, "KODE"] = e_kode
                    st.session_state["df_dapur_state"].at[
                        idx, "NAMA DAPUR"
                    ] = e_nama
                    st.session_state["df_dapur_state"].at[idx, "PIC"] = e_pic
                    st.session_state["df_dapur_state"].at[
                        idx, "ALAMAT"
                    ] = e_alamat
                    st.session_state["df_dapur_state"].at[idx, "LATITUDE"] = e_lat
                    st.session_state["df_dapur_state"].at[idx, "LONGITUDE"] = e_lon
                    st.success("Data Dapur Berhasil Diperbarui!")
                    st.rerun()

                if btn_delete:
                    st.session_state["df_dapur_state"] = st.session_state[
                        "df_dapur_state"
                    ][
                        st.session_state["df_dapur_state"]["NAMA DAPUR"]
                        != dapur_pilihan
                    ]
                    st.success("Data Dapur Berhasil Dihapus!")
                    st.rerun()

# ---------------------------------------------------------
# MODUL LAINNYA (DASHBOARD, MASTER BB, SUPPLIER, WA)
# ---------------------------------------------------------
elif menu == "📊 Dashboard & HET":
    st.subheader("📊 Dashboard Utama & Rekapitulasi Penawaran")
    m1, m2, m3 = st.columns(3)
    m1.metric("Master Item BB", len(st.session_state["df_master_bb"]))
    m2.metric("Total Dapur SPPG", len(st.session_state["df_dapur_state"]))
    m3.metric("Total Supplier", len(st.session_state["df_supplier_state"]))
    st.dataframe(st.session_state["df_master_bb"], use_container_width=True)

elif menu == "📦 Master Bahan Baku (BB)":
    st.subheader("📦 Kelola Master Data Bahan Baku (BB)")
    st.dataframe(st.session_state["df_master_bb"], use_container_width=True)

elif menu == "🤝 Data Supplier & Link Form":
    st.subheader("🤝 Kelola Data Supplier & Link Web App")
    st.dataframe(st.session_state["df_supplier_state"], use_container_width=True)

elif menu == "💬 WA & PO Generator":
    st.subheader("💬 WA Message Generator")
    st.info("Pilih supplier pada tabel untuk membuat link WhatsApp.")

elif menu == "🎯 Scoring & Evaluasi":
    st.subheader("🎯 Scoring & Evaluasi Kinerja Supplier")
    st.info("Modul Scoring Supplier SPPG.")
