import io
import os
import random
import string
import urllib.parse
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
            data_dict = {
                sheet: pd.read_excel(file_name, sheet_name=sheet)
                for sheet in xls.sheet_names
            }
            return data_dict, file_name
        except Exception as e:
            return None, str(e)
    return None, "File data lokal tidak ditemukan (Menggunakan State Default)."


data_excel, file_status = load_data()

# ---------------------------------------------------------
# 3. INITIALIZE SESSION STATES
# ---------------------------------------------------------
# Master Item Bahan Baku
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

# State Dapur
if "df_dapur_state" not in st.session_state:
    st.session_state["df_dapur_state"] = pd.DataFrame(
        {
            "KODE": ["PAKEM", "NGPLK", "SLMN4"],
            "NAMA DAPUR": [
                "PAKEM (Hargobinangun)",
                "NGEMPLAK (Umbulmartani 1)",
                "SLEMAN 4 (Triharjo)",
            ],
            "ALAMAT": [
                "Jl. Kaliurang Km 22",
                "Jl. Kaliurang Km 15.5",
                "Jl. Letkol Subadri",
            ],
            "PIC": ["SHINTA", "SHINTA", "CAHYO"],
            "LATITUDE": [-7.618382, -7.675789, -7.700475],
            "LONGITUDE": [110.426078, 110.417100, 110.342646],
            "KOTA/KABUPATEN": ["SLEMAN", "SLEMAN", "SLEMAN"],
        }
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
# 4. HANDLER FORM EXTERNAL VENDOR (DENGAN SAFE FALLBACK)
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

    # Filtering Kategori dengan Safe Fallback
    if kategori_diterima != "ALL" and str(kategori_diterima).strip() != "":
        list_kat = [
            k.replace("_", " ").strip().lower()
            for k in kategori_diterima.split(",")
        ]
        df_filtered = df_master[
            df_master["KATEGORI"].str.lower().isin(list_kat)
        ]

        # FALLBACK: Jika hasil filter kosong, tampilkan semua master BB
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
# 5. SIDEBAR NAVIGASI
# ---------------------------------------------------------
with st.sidebar:
    st.markdown("### ⚡ **SPPG Engine**")
    st.caption("Koperasi YK")
    st.markdown("---")
    st.subheader("🚀 Modul Operasional")

    modul_options = [
        "📊 Dashboard & HET",
        "📦 Master Bahan Baku (BB)",  # Modul Baru Ditambahkan
        "🏬 Kelola Data Dapur",
        "🤝 Data Supplier & Link Form",
        "💬 WA & PO Generator",
        "🎯 Scoring & Evaluasi",
    ]

    menu = st.radio("Pilih Modul:", modul_options, index=0)
    st.markdown("---")
    st.caption(f"Status Data: 🟢 {file_status}")

# ---------------------------------------------------------
# 6. HEADER UTAMA
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
# MODUL 1: DASHBOARD
# ---------------------------------------------------------
if menu == "📊 Dashboard & HET":
    st.subheader("📊 Dashboard Utama & Rekapitulasi Harga Supplier")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Master Item BB", len(st.session_state["df_master_bb"]))
    m2.metric("Total Dapur SPPG", len(st.session_state["df_dapur_state"]))
    m3.metric("Total Supplier", len(st.session_state["df_supplier_state"]))
    m4.metric(
        "Penawaran Masuk",
        len(st.session_state["df_harga_bb"]["KODE SUPPLIER"].unique()),
    )

    st.markdown("---")
    st.markdown("##### 📋 **Matriks Komparasi Penawaran Vendor vs HET**")

    df_harga = st.session_state["df_harga_bb"]
    df_master = st.session_state["df_master_bb"]

    if not df_harga.empty:
        pivot_df = df_harga.pivot(
            index="KODE_BB", columns="KODE SUPPLIER", values="HARGA PER SATUAN"
        ).reset_index()
        matrix_df = pd.merge(
            df_master,
            pivot_df,
            left_on="KODE",
            right_on="KODE_BB",
            how="left",
        ).drop(columns=["KODE_BB"], errors="ignore")
        st.dataframe(matrix_df.fillna("-"), use_container_width=True)
    else:
        st.info("Belum ada penawaran harga yang masuk dari vendor.")
        st.dataframe(df_master, use_container_width=True)

# ---------------------------------------------------------
# MODUL 2: MASTER BAHAN BAKU (BARU)
# ---------------------------------------------------------
elif menu == "📦 Master Bahan Baku (BB)":
    st.subheader("📦 Kelola Master Data Bahan Baku (BB)")

    with st.form("form_add_bb"):
        st.write("➕ **Tambah Item Bahan Baku Baru**")
        c1, c2, c3, c4 = st.columns([2, 3, 2, 2])
        with c1:
            kode_new = st.text_input(
                "Kode BB",
                value=f"KOP-BB-{len(st.session_state['df_master_bb'])+1:04d}",
            )
        with c2:
            item_new = st.text_input("Nama Item Bahan Baku")
        with c3:
            kat_new = st.selectbox(
                "Kategori",
                [
                    "Ayam",
                    "Beras",
                    "Buah",
                    "Daging",
                    "Ikan",
                    "Minyak",
                    "Sayur",
                    "Sembako",
                    "Telur Ayam",
                ],
            )
        with c4:
            sat_new = st.selectbox(
                "Satuan", ["Kg", "Liter", "Pcs", "Ikat", "Dus", "Pack"]
            )

        het_new = st.number_input("Harga HET / Acuan (Rp)", min_value=0, step=500)

        if st.form_submit_button("✨ Simpan Master BB"):
            if item_new:
                new_row = pd.DataFrame(
                    [
                        {
                            "KODE": kode_new,
                            "KATEGORI": kat_new,
                            "ITEM BB": item_new,
                            "SATUAN": sat_new,
                            "HET": het_new,
                        }
                    ]
                )
                st.session_state["df_master_bb"] = pd.concat(
                    [st.session_state["df_master_bb"], new_row],
                    ignore_index=True,
                )
                st.success("Item Master BB berhasil ditambahkan!")
                st.rerun()
            else:
                st.warning("Nama item tidak boleh kosong.")

    st.markdown("---")
    st.markdown("##### 📋 **Daftar Master Item Bahan Baku**")
    st.dataframe(st.session_state["df_master_bb"], use_container_width=True)

# ---------------------------------------------------------
# MODUL 3: KELOLA DATA DAPUR
# ---------------------------------------------------------
elif menu == "🏬 Kelola Data Dapur":
    st.subheader("🏬 Data Dapur SPPG Operasional")
    st.dataframe(st.session_state["df_dapur_state"], use_container_width=True)

# ---------------------------------------------------------
# MODUL 4: DATA SUPPLIER & LINK FORM
# ---------------------------------------------------------
elif menu == "🤝 Data Supplier & Link Form":
    st.subheader("🤝 Kelola Data Supplier & Link Web App")
    st.dataframe(
        st.session_state["df_supplier_state"][
            [
                "KODE SUPPLIER",
                "NAMA SUPPLIER",
                "NO WA",
                "JENIS BB 1",
                "LINK FORM",
            ]
        ],
        use_container_width=True,
    )

# ---------------------------------------------------------
# MODUL 5: WA & PO GENERATOR
# ---------------------------------------------------------
elif menu == "💬 WA & PO Generator":
    st.subheader("💬 WA Message Generator")
    df_sup = st.session_state["df_supplier_state"]

    if not df_sup.empty:
        selected_sup = st.selectbox(
            "Pilih Supplier Target:", df_sup["NAMA SUPPLIER"].tolist()
        )
        row_sup = df_sup[df_sup["NAMA SUPPLIER"] == selected_sup].iloc[0]

        link_form = (
            f"{WEB_APP_URL}?token={row_sup['TOKEN']}&kat={row_sup['JENIS BB 1']}"
        )
        msg = (
            f"Halo *{row_sup['NAMA SUPPLIER']}*,\nMohon update harga penawaran"
            f" melalui link:\n{link_form}"
        )

        st.text_area("Draft Pesan WA:", msg, height=120)
        wa_url = f"https://wa.me/{row_sup['NO WA']}?text={urllib.parse.quote(msg)}"
        st.markdown(f"[📲 Kirim WhatsApp ke Vendor]({wa_url})")

# ---------------------------------------------------------
# MODUL 6: SCORING
# ---------------------------------------------------------
elif menu == "🎯 Scoring & Evaluasi":
    st.subheader("🎯 Scoring & Evaluasi Kinerja Supplier")
