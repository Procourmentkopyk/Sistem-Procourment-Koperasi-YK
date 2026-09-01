import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import random
import string
import urllib.parse
from math import radians, cos, sin, asin, sqrt

# =========================================================
# KONFIGURASI
# =========================================================
st.set_page_config(
    page_title="SPPG Procurement Engine | Koperasi YK",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

WEB_APP_URL = "https://sistem-procurement-koperasi-yk.streamlit.app"
MASTER_FILE = "UPDATE Harga Procourment.xlsx"

SHEET_HARGA = "UPDATE HARGA BB"
SHEET_JENIS = "Supplier Jenis BB"
SHEET_LINK = "Supplier Link"

# Baris Excel:
# baris 1 = periode / judul
# baris 2 = nomor/index supplier
# baris 3 = header utama
EXCEL_HEADER_ROW = 2

# Kolom internal yang boleh dilihat admin, tetapi TIDAK BOLEH
# ditampilkan pada form supplier.
INTERNAL_ONLY_COLUMNS = {
    "TARGET HARGA",
    "HET",
    "HET PATOKAN",
    "HARGA PASAR",
    "HARGA PASAR RATA",
    "SKOR",
    "RANKING",
}

# =========================================================
# CSS
# =========================================================
st.markdown(
    """
    <style>
    .header-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        padding: 1.25rem 1.5rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,.10);
    }
    .supplier-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 18px;
        margin-bottom: 15px;
    }
    .notice {
        background: #eff6ff;
        border-left: 5px solid #2563eb;
        padding: 12px 15px;
        border-radius: 8px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# UTILITY
# =========================================================
def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def buat_token(length=32):
    return "".join(
        random.choices(string.ascii_letters + string.digits, k=length)
    )


def normalize_category(value):
    """Normalisasi kategori agar OLAHAN dan Olahan dianggap sama."""
    text = clean_text(value).upper()
    return " ".join(text.split())


def normalize_code(value):
    return clean_text(value).upper()


def rupiah(value):
    if pd.isna(value) or value in ("", None):
        return "-"
    try:
        return f"Rp {float(value):,.0f}"
    except Exception:
        return str(value)


def hitung_jarak_haversine(lat1, lon1, lat2, lon2):
    if any(pd.isna(x) for x in [lat1, lon1, lat2, lon2]):
        return np.nan

    R = 6371.0
    lat1, lon1, lat2, lon2 = map(
        radians, [float(lat1), float(lon1), float(lat2), float(lon2)]
    )
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = (
        sin(dlat / 2) ** 2
        + cos(lat1) * cos(lat2) * sin(dlon / 2) ** 2
    )
    c = 2 * asin(sqrt(a))
    return R * c


# =========================================================
# EXPORT HELPER
# =========================================================
def _excel_bytes(df, sheet_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


# =========================================================
# LOAD EXCEL MASTER
# =========================================================
@st.cache_data
def load_master_excel(file_name=MASTER_FILE):
    if not os.path.exists(file_name):
        return None, f"File '{file_name}' tidak ditemukan."

    try:
        # UPDATE HARGA BB harus dibaca header utama pada baris ke-3.
        df_harga = pd.read_excel(
            file_name,
            sheet_name=SHEET_HARGA,
            header=EXCEL_HEADER_ROW,
        )

        df_jenis = pd.read_excel(
            file_name,
            sheet_name=SHEET_JENIS,
            header=0,
        )

        df_link = pd.read_excel(
            file_name,
            sheet_name=SHEET_LINK,
            header=0,
        )

        # Bersihkan nama kolom
        df_harga.columns = [
            clean_text(c).upper() if not pd.isna(c) else ""
            for c in df_harga.columns
        ]
        df_jenis.columns = [
            clean_text(c).upper() if not pd.isna(c) else ""
            for c in df_jenis.columns
        ]
        df_link.columns = [
            clean_text(c).upper() if not pd.isna(c) else ""
            for c in df_link.columns
        ]

        # Hapus kolom tanpa nama
        df_harga = df_harga.loc[
            :,
            [c for c in df_harga.columns if c and not c.startswith("UNNAMED")]
        ]

        # Buang baris yang tidak mempunyai item
        if "ITEM BB" in df_harga.columns:
            df_harga = df_harga[
                df_harga["ITEM BB"].notna()
                & (df_harga["ITEM BB"].astype(str).str.strip() != "")
            ].copy()

        # Normalisasi Supplier Jenis BB
        if "KODE SUPPLIER" in df_jenis.columns:
            df_jenis["KODE SUPPLIER"] = df_jenis["KODE SUPPLIER"].map(
                normalize_code
            )

        # Normalisasi Supplier Link
        if "KODE SUPPLIER" in df_link.columns:
            df_link["KODE SUPPLIER"] = df_link["KODE SUPPLIER"].map(
                normalize_code
            )

        return {
            "harga": df_harga.reset_index(drop=True),
            "jenis": df_jenis.reset_index(drop=True),
            "link": df_link.reset_index(drop=True),
        }, "OK"

    except Exception as e:
        return None, f"Gagal membaca Excel: {e}"


master, master_status = load_master_excel()

# =========================================================
# SESSION STATE
# =========================================================
if "master" not in st.session_state:
    st.session_state["master"] = master

if "master_status" not in st.session_state:
    st.session_state["master_status"] = master_status

if "supplier_links" not in st.session_state:
    if master and isinstance(master, dict):
        st.session_state["supplier_links"] = master["link"].copy()
    else:
        st.session_state["supplier_links"] = pd.DataFrame(
            columns=[
                "KODE SUPPLIER",
                "NAMA SUPPLIER",
                "TOKEN",
                "LINK FORM",
                "STATUS",
            ]
        )

if "harga_matrix" not in st.session_state:
    if master and isinstance(master, dict):
        st.session_state["harga_matrix"] = master["harga"].copy()
    else:
        st.session_state["harga_matrix"] = pd.DataFrame()

# =========================================================
# TOKEN / SUPPLIER FUNCTIONS
# =========================================================
def ensure_supplier_links():
    """
    Memastikan setiap kode supplier di UPDATE HARGA BB memiliki
    satu token permanen.

    Token lama dipertahankan.
    Supplier baru mendapat token baru.
    """
    df_harga = st.session_state["harga_matrix"].copy()
    df_link = st.session_state["supplier_links"].copy()

    if df_harga.empty:
        return df_link

    # Supplier = semua kolom setelah TARGET HARGA.
    base_cols = {
        "NO",
        "P/N",
        "KATEGORI",
        "JENIS BB",
        "ITEM BB",
        "SAT",
        "DAPUR",
        "SAT PASAR",
        "TARGET HARGA",
        "PERIODE",
    }

    supplier_cols = [
        c for c in df_harga.columns
        if clean_text(c).upper() not in base_cols
        and not clean_text(c).upper().startswith("UNNAMED")
    ]

    # Hanya kode pendek supplier seperti DNS, CTB, SD3, dst.
    supplier_cols = [
        clean_text(c).upper()
        for c in supplier_cols
        if clean_text(c)
    ]

    if "KODE SUPPLIER" not in df_link.columns:
        df_link["KODE SUPPLIER"] = ""

    if "NAMA SUPPLIER" not in df_link.columns:
        df_link["NAMA SUPPLIER"] = ""

    if "TOKEN" not in df_link.columns:
        df_link["TOKEN"] = ""

    if "LINK FORM" not in df_link.columns:
        df_link["LINK FORM"] = ""

    if "STATUS" not in df_link.columns:
        df_link["STATUS"] = ""

    existing = {
        normalize_code(row["KODE SUPPLIER"]): idx
        for idx, row in df_link.iterrows()
        if clean_text(row.get("KODE SUPPLIER", ""))
    }

    for code in supplier_cols:
        if code in existing:
            idx = existing[code]
            token = clean_text(df_link.at[idx, "TOKEN"])

            if not token:
                token = buat_token()

            df_link.at[idx, "TOKEN"] = token
            df_link.at[idx, "LINK FORM"] = (
                f"{WEB_APP_URL}/?token={urllib.parse.quote(token)}"
            )
            df_link.at[idx, "STATUS"] = "ok"

            if not clean_text(df_link.at[idx, "NAMA SUPPLIER"]):
                df_link.at[idx, "NAMA SUPPLIER"] = code

        else:
            token = buat_token()
            new_row = {
                "KODE SUPPLIER": code,
                "NAMA SUPPLIER": code,
                "TOKEN": token,
                "LINK FORM": (
                    f"{WEB_APP_URL}/?token={urllib.parse.quote(token)}"
                ),
                "STATUS": "ok",
            }
            df_link = pd.concat(
                [df_link, pd.DataFrame([new_row])],
                ignore_index=True,
            )

    st.session_state["supplier_links"] = df_link
    return df_link


def get_supplier_by_token(token):
    token = clean_text(token)

    if not token:
        return None

    df_link = ensure_supplier_links()

    if df_link.empty or "TOKEN" not in df_link.columns:
        return None

    match = df_link[
        df_link["TOKEN"].astype(str).str.strip() == token
    ]

    if match.empty:
        return None

    return match.iloc[0]


def get_supplier_categories(kode_supplier):
    """
    Mengambil maksimal 6 jenis BB dari sheet Supplier Jenis BB.
    Target/HET tidak ikut diambil.
    """
    kode_supplier = normalize_code(kode_supplier)
    df = master["jenis"] if master else pd.DataFrame()

    if df.empty or "KODE SUPPLIER" not in df.columns:
        return []

    match = df[
        df["KODE SUPPLIER"].map(normalize_code) == kode_supplier
    ]

    if match.empty:
        return []

    row = match.iloc[0]
    categories = []

    for i in range(1, 7):
        col = f"JENIS BB {i}"
        if col in row.index:
            value = clean_text(row[col])
            if value and normalize_category(value) != "END":
                categories.append(value)

    # Hilangkan duplikat
    result = []
    seen = set()

    for value in categories:
        key = normalize_category(value)
        if key not in seen:
            seen.add(key)
            result.append(value)

    return result


def get_supplier_items(kode_supplier):
    """
    Mengambil item yang kategorinya sesuai hak supplier.

    PENTING:
    kolom TARGET HARGA sengaja tidak pernah dikirim ke form supplier.
    """
    df = st.session_state["harga_matrix"].copy()

    if df.empty:
        return pd.DataFrame()

    categories = get_supplier_categories(kode_supplier)
    allowed = {normalize_category(x) for x in categories}

    if "JENIS BB" not in df.columns:
        return pd.DataFrame()

    result = df[
        df["JENIS BB"].map(normalize_category).isin(allowed)
    ].copy()

    # Tampilkan hanya kolom publik.
    public_cols = [
        c for c in [
            "NO",
            "P/N",
            "JENIS BB",
            "ITEM BB",
            "SAT",
        ]
        if c in result.columns
    ]

    return result[public_cols].copy()


def simpan_penawaran_supplier(kode_supplier, harga_dict):
    """
    Menyimpan harga supplier ke kolom supplier yang bersangkutan.
    Harga target tidak disentuh dan tidak dikirim ke supplier.
    """
    df = st.session_state["harga_matrix"].copy()

    if kode_supplier not in df.columns:
        # Coba cari case-insensitive
        found = next(
            (
                c for c in df.columns
                if normalize_code(c) == normalize_code(kode_supplier)
            ),
            None,
        )
        if found:
            kode_supplier = found
        else:
            raise ValueError(
                f"Kolom supplier '{kode_supplier}' tidak ditemukan "
                "di sheet UPDATE HARGA BB."
            )

    if "P/N" not in df.columns:
        raise ValueError("Kolom P/N tidak ditemukan di master.")

    jumlah = 0

    for pn, harga in harga_dict.items():
        if harga is None or harga == "":
            continue

        try:
            nilai = float(harga)
        except Exception:
            continue

        if nilai <= 0:
            continue

        mask = df["P/N"].astype(str).str.strip() == str(pn).strip()

        if mask.any():
            df.loc[mask, kode_supplier] = nilai
            jumlah += 1

    st.session_state["harga_matrix"] = df
    return jumlah


# =========================================================
# AUTO GENERATE / MAINTAIN TOKEN
# =========================================================
if master and isinstance(master, dict):
    ensure_supplier_links()

# =========================================================
# MODE SUPPLIER EXTERNAL
# =========================================================
query_params = st.query_params
token_diterima = clean_text(query_params.get("token", ""))

if token_diterima:
    supplier_info = get_supplier_by_token(token_diterima)

    st.markdown(
        """
        <div class="header-card">
            <h2 style="margin:0;color:white;">📝 Form Penawaran Harga</h2>
            <p style="margin:5px 0 0;color:#cbd5e1;">
                Koperasi YK — Sistem Procurement SPPG
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if supplier_info is None:
        st.error(
            "❌ Link supplier tidak valid atau token tidak ditemukan."
        )
        st.stop()

    kode_supplier = normalize_code(
        supplier_info.get("KODE SUPPLIER", "")
    )
    nama_supplier = clean_text(
        supplier_info.get("NAMA SUPPLIER", "")
    ) or kode_supplier

    categories = get_supplier_categories(kode_supplier)
    items = get_supplier_items(kode_supplier)

    st.success(
        f"Selamat datang, **{nama_supplier}** "
        f"(`{kode_supplier}`)"
    )

    if categories:
        st.info(
            "Kategori yang dapat Anda tawarkan: "
            + ", ".join(categories)
        )
    else:
        st.warning(
            "Belum ada jenis BB yang ditetapkan untuk supplier ini."
        )

    # =====================================================
    # PENTING:
    # Tidak ada TARGET HARGA, HET, harga pasar, atau
    # harga supplier lain pada mode external.
    # =====================================================
    st.markdown(
        """
        <div class="notice">
        🔒 <b>Kerahasiaan Penawaran</b><br>
        Harga yang Anda masukkan hanya merupakan penawaran
        dari supplier Anda. Sistem tidak menampilkan target harga
        maupun penawaran supplier lain.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("---")

    if items.empty:
        st.warning(
            "Belum ada item bahan baku yang sesuai kategori supplier."
        )
        st.stop()

    search = st.text_input(
        "🔍 Cari bahan baku",
        placeholder="Contoh: telur, ayam, beras...",
        key="supplier_search",
    )

    items_display = items.copy()

    if search:
        text = search.strip()
        mask = (
            items_display["ITEM BB"]
            .astype(str)
            .str.contains(text, case=False, na=False)
            |
            items_display["JENIS BB"]
            .astype(str)
            .str.contains(text, case=False, na=False)
        )
        items_display = items_display[mask]

    # Group berdasarkan jenis BB agar supplier lebih mudah mengisi.
    with st.form("form_penawaran_supplier"):
        st.subheader("🛒 Masukkan Harga Penawaran")

        harga_input = {}

        for kategori, group in items_display.groupby(
            "JENIS BB", sort=False
        ):
            st.markdown(f"### 📦 {kategori}")

            for _, row in group.iterrows():
                pn = clean_text(row.get("P/N", ""))
                item = clean_text(row.get("ITEM BB", ""))
                sat = clean_text(row.get("SAT", ""))

                col1, col2, col3 = st.columns([4, 1, 2])

                with col1:
                    st.write(f"**{item}**")
                    st.caption(pn)

                with col2:
                    st.write(sat or "-")

                with col3:
                    harga_input[pn] = st.number_input(
                        "Harga",
                        min_value=0,
                        step=500,
                        value=0,
                        key=f"harga_{kode_supplier}_{pn}",
                        label_visibility="collapsed",
                    )

        submit = st.form_submit_button(
            "🚀 KIRIM PENAWARAN HARGA",
            type="primary",
            use_container_width=True,
        )

    if submit:
        jumlah = simpan_penawaran_supplier(
            kode_supplier,
            harga_input,
        )

        if jumlah == 0:
            st.warning(
                "Belum ada harga yang diisi. "
                "Masukkan minimal satu harga."
            )
        else:
            st.balloons()
            st.success(
                f"✅ {jumlah} penawaran berhasil diterima "
                f"dari supplier **{nama_supplier}**."
            )

            st.info(
                "Penawaran telah masuk ke sistem. "
                "Anda dapat menutup halaman ini."
            )

    st.stop()

# =========================================================
# SIDEBAR ADMIN
# =========================================================
with st.sidebar:
    st.markdown("### ⚡ **SPPG Engine**")
    st.caption("Koperasi YK")
    st.markdown("---")

    st.subheader("🚀 Modul Operasional")

    menu = st.radio(
        "Pilih Modul:",
        [
            "📊 Dashboard & HET",
            "🛒 Update Harga BB",
            "🤝 Data Supplier & Link Form",
            "📈 Komparasi Harga",
            "🎯 Scoring & Evaluasi",
        ],
        index=0,
    )

    st.markdown("---")
    st.caption(
        f"Status Master: {'🟢 Terbaca' if master else '🔴 Error'}"
    )

# =========================================================
# HEADER ADMIN
# =========================================================
st.markdown(
    """
    <div class="header-card">
        <h2 style="margin:0;color:white;">
            🏭 Enterprise Procurement System
        </h2>
        <p style="margin:0;color:#cbd5e1;">
            Koperasi YK — Sistem Evaluasi Supplier & Dapur SPPG
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

if not master:
    st.error(st.session_state["master_status"])
    st.info(
        f"Pastikan file **{MASTER_FILE}** berada di folder yang sama "
        "dengan app.py pada repository GitHub."
    )
    st.stop()

df_harga = st.session_state["harga_matrix"]
df_link = st.session_state["supplier_links"]

# =========================================================
# DASHBOARD
# =========================================================
if menu == "📊 Dashboard & HET":
    st.subheader("📊 Dashboard Utama")

    supplier_cols = [
        c for c in df_harga.columns
        if c not in {
            "NO", "P/N", "KATEGORI", "JENIS BB",
            "ITEM BB", "SAT", "DAPUR", "SAT PASAR",
            "TARGET HARGA", "PERIODE"
        }
        and clean_text(c)
    ]

    total_supplier = len(supplier_cols)
    total_item = (
        df_harga["ITEM BB"].nunique()
        if "ITEM BB" in df_harga.columns
        else 0
    )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Supplier", total_supplier)
    m2.metric("Total Item BB", total_item)
    m3.metric("Supplier Link", len(df_link))
    m4.metric("Target Harga", "🔒 Internal")

    st.markdown("---")

    st.info(
        "Target harga, HET, harga pasar, dan data supplier lain "
        "hanya ditampilkan pada sisi admin. Data tersebut tidak "
        "pernah dikirim ke form supplier."
    )

# =========================================================
# UPDATE HARGA BB
# =========================================================
elif menu == "🛒 Update Harga BB":
    st.subheader("🛒 Update Harga Bahan Baku")

    if df_harga.empty:
        st.warning("Data harga kosong.")
    else:
        search = st.text_input(
            "🔍 Cari item / P/N / kategori / supplier"
        )

        df_view = df_harga.copy()

        if search:
            mask = pd.Series(False, index=df_view.index)

            for col in [
                "P/N",
                "KATEGORI",
                "JENIS BB",
                "ITEM BB",
            ]:
                if col in df_view.columns:
                    mask |= df_view[col].astype(str).str.contains(
                        search,
                        case=False,
                        na=False,
                    )

            df_view = df_view[mask]

        st.caption(
            f"Menampilkan {len(df_view)} dari {len(df_harga)} item."
        )

        st.dataframe(
            df_view,
            use_container_width=True,
            height=600,
        )

        st.download_button(
            "📥 Export Master Harga",
            data=_excel_bytes(df_harga, "UPDATE HARGA BB"),
            file_name="UPDATE_HARGA_BB_KOPERASI_YK.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

# =========================================================
# SUPPLIER & LINK FORM
# =========================================================
elif menu == "🤝 Data Supplier & Link Form":
    st.subheader("🤝 Supplier & Link Form")

    # Admin boleh melihat token dan link.
    # Supplier tidak pernah mendapatkan tabel ini.
    st.dataframe(
        df_link[
            [
                c for c in [
                    "KODE SUPPLIER",
                    "NAMA SUPPLIER",
                    "TOKEN",
                    "LINK FORM",
                    "STATUS",
                ]
                if c in df_link.columns
            ]
        ],
        use_container_width=True,
        height=600,
    )

    st.info(
        "🔐 Link pada tabel ini adalah link pribadi supplier. "
        "Jangan membagikan link supplier A kepada supplier B."
    )

# =========================================================
# KOMPARASI HARGA
# =========================================================
elif menu == "📈 Komparasi Harga":
    st.subheader("📈 Komparasi Penawaran Supplier")

    supplier_cols = [
        c for c in df_harga.columns
        if c not in {
            "NO", "P/N", "KATEGORI", "JENIS BB",
            "ITEM BB", "SAT", "DAPUR", "SAT PASAR",
            "TARGET HARGA", "PERIODE"
        }
        and clean_text(c)
    ]

    selected_item = st.selectbox(
        "Pilih bahan baku",
        df_harga["ITEM BB"].dropna().astype(str).unique()
        if "ITEM BB" in df_harga.columns
        else [],
    )

    if selected_item:
        row = df_harga[
            df_harga["ITEM BB"].astype(str) == selected_item
        ].iloc[0]

        rows = []

        for supplier in supplier_cols:
            nilai = pd.to_numeric(
                pd.Series([row.get(supplier)]),
                errors="coerce",
            ).iloc[0]

            if pd.notna(nilai) and float(nilai) > 0:
                rows.append({
                    "SUPPLIER": supplier,
                    "HARGA PENAWARAN": float(nilai),
                })

        result = pd.DataFrame(rows)

        if result.empty:
            st.info(
                "Belum ada supplier yang mengirim harga untuk item ini."
            )
        else:
            # Target hanya tampil di sisi admin.
            target = pd.to_numeric(
                pd.Series([row.get("TARGET HARGA")]),
                errors="coerce",
            ).iloc[0]

            if pd.notna(target):
                result["TARGET HARGA"] = float(target)
                result["SELISIH TARGET"] = (
                    result["HARGA PENAWARAN"] - float(target)
                )
                result["STATUS"] = np.where(
                    result["SELISIH TARGET"] <= 0,
                    "🟢 Di bawah/sesuai target",
                    "🔴 Di atas target",
                )

            result = result.sort_values(
                "HARGA PENAWARAN"
            ).reset_index(drop=True)

            st.dataframe(
                result,
                use_container_width=True,
            )

# =========================================================
# SCORING
# =========================================================
elif menu == "🎯 Scoring & Evaluasi":
    st.subheader("🎯 Scoring & Evaluasi Supplier")

    w_harga = st.slider(
        "Bobot Harga (%)",
        0,
        100,
        40,
    )

    w_jarak = st.slider(
        "Bobot Jarak (%)",
        0,
        100 - w_harga,
        min(30, 100 - w_harga),
    )

    w_sla = 100 - w_harga - w_jarak

    st.info(
        f"Bobot SLA otomatis: **{w_sla}%**"
    )

    st.markdown("---")
    st.warning(
        "Versi ini sudah menyiapkan struktur scoring. "
        "Nilai jarak dan SLA sebaiknya diambil dari data operasional "
        "sebenarnya, bukan angka simulasi."
    )
