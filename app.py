import io
import os
import random
import string
import urllib.parse
import numpy as np
import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# 1. KONFIGURASI HALAMAN
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


# CSS Header & Style
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
# 2. INITIALIZE SESSION STATE
# ---------------------------------------------------------
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

if "df_supplier_state" not in st.session_state:
    st.session_state["df_supplier_state"] = pd.DataFrame(
        [
            {
                "KODE SUPPLIER": "DNS",
                "NAMA SUPPLIER": "UD Tani Makmur",
                "NO WA": "6281234567890",
                "JENIS BB 1": "Ayam",
                "JENIS BB 2": "Telur Ayam",
                "JENIS BB 3": "-",
                "TOKEN": "sample1",
                "LINK FORM": f"{WEB_APP_URL}?token=sample1",
                "PIC": "Budi",
                "ALAMAT": "Sleman",
            }
        ]
    )

# Master Item Bahan Baku (Menyamakan Struktur Baris Spreadsheet)
if "df_master_bb" not in st.session_state:
    st.session_state["df_master_bb"] = pd.DataFrame(
        [
            {
                "KODE": "KOP-BB-0001",
                "KATEGORI": "Karbohidrat",
                "JENIS BB": "Beras",
                "ITEM BB": "Beras Premium",
                "SATUAN": "Kg",
                "TARGET HARGA": 14900,
            },
            {
                "KODE": "KOP-BB-0002",
                "KATEGORI": "Karbohidrat",
                "JENIS BB": "Beras",
                "ITEM BB": "Beras Medium",
                "SATUAN": "Kg",
                "TARGET HARGA": 13300,
            },
            {
                "KODE": "KOP-BB-0006",
                "KATEGORI": "Protein Hewani",
                "JENIS BB": "Telur Ayam",
                "ITEM BB": "Telur Ayam",
                "SATUAN": "Kg",
                "TARGET HARGA": 25500,
            },
            {
                "KODE": "KOP-BB-0014",
                "KATEGORI": "Protein Hewani",
                "JENIS BB": "Ayam",
                "ITEM BB": "Ayam Karkas 1kg = 8 potong",
                "SATUAN": "Kg",
                "TARGET HARGA": 43000,
            },
        ]
    )

# Tabel penampung relational input harga dari supplier
if "df_input_supplier" not in st.session_state:
    st.session_state["df_input_supplier"] = (
        pd.DataFrame()
    )  # Col: KODE_SUPPLIER, KODE_BB, HARGA

# ---------------------------------------------------------
# 3. HANDLER UNTUK FORM EXTERNAL VENDOR (MATRIX INPUT)
# ---------------------------------------------------------
query_params = st.query_params

if "token" in query_params:
    token_diterima = query_params["token"]

    # Cari Kode Supplier berdasarkan Token
    df_sup = st.session_state["df_supplier_state"]
    sup_match = df_sup[df_sup["TOKEN"] == token_diterima]

    if not sup_match.empty:
        kode_sup = sup_match.iloc[0]["KODE SUPPLIER"]
        nama_sup = sup_match.iloc[0]["NAMA SUPPLIER"]
    else:
        kode_sup = token_diterima.upper()
        nama_sup = "Vendor External"

    st.title("📝 Form Penawaran Harga Supplier")
    st.caption(f"Supplier: **{kode_sup} - {nama_sup}**")
    st.markdown("---")

    df_master = st.session_state["df_master_bb"]

    # Tombol Kembali Khusus Jika Dibuka oleh Admin
    if st.sidebar.button("⬅️ Kembali ke Dashboard Admin"):
        st.query_params.clear()
        st.rerun()

    with st.form("form_matrix_supplier"):
        st.subheader("🛒 Isikan Harga Penawaran Sesuai Item")
        input_harga = {}

        # Render form pengisian harga berdasarkan list item
        for idx, row in df_master.iterrows():
            c1, c2, c3 = st.columns([3, 1, 2])
            with c1:
                st.write(
                    f"**{row['ITEM BB']}** (`{row['KODE']}`)\n*{row['KATEGORI']} - {row['JENIS BB']}*"
                )
            with c2:
                st.write(f"Satuan: **{row['SATUAN']}**")
            with c3:
                input_harga[row["KODE"]] = st.number_input(
                    "Harga (Rp)",
                    min_value=0,
                    step=500,
                    value=0,
                    key=f"input_{row['KODE']}",
                )
            st.divider()

        if st.form_submit_button(
            "🚀 Kirim Penawaran Harga", use_container_width=True
        ):
            new_records = []
            for kode_bb, harga in input_harga.items():
                if harga > 0:
                    new_records.append(
                        {
                            "KODE_SUPPLIER": kode_sup,
                            "KODE_BB": kode_bb,
                            "HARGA": harga,
                        }
                    )

            if new_records:
                df_new = pd.DataFrame(new_records)
                # Overwrite jika supplier tersebut pernah mengisi
                if not st.session_state["df_input_supplier"].empty:
                    st.session_state["df_input_supplier"] = st.session_state[
                        "df_input_supplier"
                    ][
                        st.session_state["df_input_supplier"]["KODE_SUPPLIER"]
                        != kode_sup
                    ]

                st.session_state["df_input_supplier"] = pd.concat(
                    [st.session_state["df_input_supplier"], df_new],
                    ignore_index=True,
                )
                st.balloons()
                st.success("Penawaran harga berhasil disimpan!")
            else:
                st.warning("Mohon isi minimal satu harga item.")

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
        "📊 Dashboard & Rekap Matriks",
        "🏬 Kelola Data Dapur",
        "🤝 Data Supplier & Link Form",
        "💬 WA & PO Generator",
        "🚛 Matriks Jarak",
        "🎯 Scoring & Evaluasi",
        "📈 HET & Komparasi Pasar",
    ]

    menu = st.radio("Pilih Modul:", modul_options, index=0)
    st.markdown("---")
    st.caption("Status Data: 🟢 Sistem Evaluasi Supplier & Dapur.xlsb")

# HEADER UTAMA
st.markdown(
    """
<div class="header-card">
    <div>
        <h2 style="margin:0; color: white;">🏭 Enterprise Procurement System</h2>
        <p style="margin:0; opacity: 0.8;">Koperasi YK — Sistem Evaluasi Supplier & Dapur SPPG</p>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------
# 5. RENDER MODUL BERDASARKAN PILIHAN MENU
# ---------------------------------------------------------
if menu == "📊 Dashboard & Rekap Matriks":
    st.subheader("📊 Dashboard & Rekapitulasi Matriks Harga")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Dapur SPPG", len(st.session_state["df_dapur_state"]))
    m2.metric("Total Supplier", len(st.session_state["df_supplier_state"]))
    m3.metric("Item Bahan Baku", len(st.session_state["df_master_bb"]))
    m4.metric("Rata-rata Ketepatan", "94.2%")

    st.markdown("---")
    st.subheader("📋 Matriks Update Harga Procurement")

    df_master = st.session_state["df_master_bb"]
    df_input = st.session_state["df_input_supplier"]

    if not df_input.empty:
        # PIVOT TABEL: Menjadikan Kode Supplier sebagai Kolom
        pivot_df = df_input.pivot(
            index="KODE_BB", columns="KODE_SUPPLIER", values="HARGA"
        ).reset_index()

        # Merge dengan Master Bahan Baku
        final_matrix = pd.merge(
            df_master,
            pivot_df,
            left_on="KODE",
            right_on="KODE_BB",
            how="left",
        ).drop(columns=["KODE_BB"], errors="ignore")
        final_matrix = final_matrix.fillna("-")

        st.dataframe(final_matrix, use_container_width=True)

        # DOWNLOAD BUTTON EXCEL
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            final_matrix.to_excel(
                writer, sheet_name="UPDATE HARGA BB", index=False
            )

        st.download_button(
            label="📥 Download Hasil Matriks (.xlsx)",
            data=buffer.getvalue(),
            file_name="UPDATE_Harga_Procurement.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    else:
        st.info("Belum ada data penawaran yang masuk dari supplier.")
        st.dataframe(df_master, use_container_width=True)

elif menu == "🏬 Kelola Data Dapur":
    st.subheader("🏬 Manajemen Data Dapur SPPG & Peta Sebaran")
    st.dataframe(st.session_state["df_dapur_state"], use_container_width=True)

elif menu == "🤝 Data Supplier & Link Form":
    st.subheader("🤝 Kelola Data Supplier & Link Web App")
    st.dataframe(st.session_state["df_supplier_state"], use_container_width=True)

elif menu == "💬 WA & PO Generator":
    st.subheader("💬 WhatsApp Message & PO Link Generator")

    df_sup = st.session_state["df_supplier_state"]
    list_supplier_display = [
        f"{r['KODE SUPPLIER']} - {r['NAMA SUPPLIER']}"
        for _, r in df_sup.iterrows()
    ]
    selected_sup = st.selectbox(
        "Pilih Supplier Target:", list_supplier_display
    )

    idx_sup = list_supplier_display.index(selected_sup)
    row_sup = df_sup.iloc[idx_sup]

    LIST_KATEGORI = [
        "Ayam",
        "Beras",
        "Buah",
        "Daging",
        "Ikan",
        "Sembako",
        "Telur Ayam",
    ]
    selected_kategori = st.multiselect(
        "Pilih Kategori Barang:", options=LIST_KATEGORI, default=["Ayam"]
    )

    token_sup = row_sup.get("TOKEN", "sample1")
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
        st.markdown(
            f'<a href="{wa_url}" target="_blank"><button style="background-color:#25D366; color:white; border:none; padding:10px; border-radius:6px; cursor:pointer; width:100%;">💬 Kirim via WhatsApp Direct</button></a>',
            unsafe_allow_html=True,
        )

elif menu == "🚛 Matriks Jarak":
    st.subheader("🚛 Matriks Jarak & Rute Distribusi Dapur SPPG")

elif menu == "🎯 Scoring & Evaluasi":
    st.subheader("🎯 Scoring & Evaluasi Kinerja Supplier")

elif menu == "📈 HET & Komparasi Pasar":
    st.subheader("📈 Pemantauan Harga HET & Komparasi Pasar")
