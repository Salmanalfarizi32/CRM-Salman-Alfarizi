import streamlit as st
import pandas as pd
from pathlib import Path

st.set_page_config(page_title="CRM Dashboard - Excel Auto Mode", layout="wide")
st.title("📊 CRM Dashboard (Auto Excel Mode)")

# Lokasi file Excel
EXCEL_PATH = Path("data/CRM Salman Alfarizi.xlsx")

# Daftar sheet yang diharapkan
SHEET_NAMES = [
    "VIP BUYER",
    "Kategori Buyer",
    "Marketing_ads",
    "Pertumbuhan Pelanggan",
    "Produk Populer",
    "Produk Favorit Customer"
]

# Fungsi untuk load semua sheet
@st.cache_data(ttl=3600)
def load_excel(path):
    return pd.read_excel(path, sheet_name=None)

# Cek apakah file ada
if not EXCEL_PATH.exists():
    st.error(f"❌ File tidak ditemukan: {EXCEL_PATH}. Pastikan file sudah di-commit ke repo.")
else:
    try:
        # Load semua sheet
        all_sheets = load_excel(EXCEL_PATH)
        available = [s for s in SHEET_NAMES if s in all_sheets.keys()]
        missing = [s for s in SHEET_NAMES if s not in all_sheets.keys()]

        if missing:
            st.warning(f"⚠️ Sheet tidak ditemukan di file: {', '.join(missing)}")

        # Pilihan sheet di sidebar
        selected = st.sidebar.selectbox("Pilih sheet", available)
        df = all_sheets[selected]

        # Format angka otomatis jadi rupiah kalau cocok
        for col in df.columns:
            if df[col].dtype in ["int64", "float64"] and "ID" not in col:
                df[col] = df[col].apply(lambda x: f"Rp {x:,.0f}".replace(",", ".") if pd.notnull(x) else "-")

        # Tampilkan tabel rapi
        st.subheader(f"📄 {selected}")
        st.dataframe(df, use_container_width=True)

    except Exception as e:
        st.error(f"Error baca Excel: {e}")
