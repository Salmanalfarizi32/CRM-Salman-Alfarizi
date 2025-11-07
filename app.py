import streamlit as st
import pandas as pd
from pathlib import Path

st.set_page_config(page_title="CRM Dashboard - Excel Auto Mode", layout="wide")
st.title("📊 CRM Dashboard (Auto Excel Mode)")

# Lokasi file di repo (misal taruh di folder data/)
EXCEL_PATH = Path("data/CRM Salman Alfarizi.xlsx")

SHEET_NAMES = [
    "VIP BUYER",
    "Kategori Buyer",
    "Marketing_ads",
    "Pertumbuhan Pelanggan",
    "Produk Populer",
    "Produk Favorit Customer"
]

@st.cache_data(ttl=3600)
def load_excel(path):
    return pd.read_excel(path, sheet_name=None, engine="openpyxl")

if not EXCEL_PATH.exists():
    st.error(f"❌ File tidak ditemukan: {EXCEL_PATH}. Pastikan file sudah di-commit ke repo.")
else:
    try:
        all_sheets = load_excel(EXCEL_PATH)
        available = [s for s in SHEET_NAMES if s in all_sheets.keys()]
        missing = [s for s in SHEET_NAMES if s not in all_sheets.keys()]

        if missing:
            st.warning(f"⚠️ Sheet tidak ditemukan di file: {', '.join(missing)}")

        selected = st.sidebar.selectbox("Pilih sheet", available)
        df = all_sheets[selected]

        st.subheader(f"📄 {selected}")
        st.write(f"Rows: {df.shape[0]} — Columns: {df.shape[1]}")
        st.dataframe(df, use_container_width=True)

    except Exception as e:
        st.error(f"Error baca Excel: {e}")
