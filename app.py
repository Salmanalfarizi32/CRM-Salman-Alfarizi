import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="CRM Dashboard - Excel Mode", layout="wide")
st.title("📊 CRM Dashboard (Excel Mode)")

uploaded = st.file_uploader("Upload file Excel kamu (.xlsx)", type=["xlsx"])

SHEET_NAMES = [
    "VIP BUYER",
    "Kategori Buyer",
    "Marketing_ads",
    "Pertumbuhan Pelanggan",
    "Produk Populer",
    "Produk Favorit Customer"
]

@st.cache_data(ttl=3600)
def load_excel(file_bytes):
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, engine="openpyxl")

if uploaded:
    try:
        all_sheets = load_excel(uploaded.getvalue())
        available = [s for s in SHEET_NAMES if s in all_sheets.keys()]
        missing = [s for s in SHEET_NAMES if s not in all_sheets.keys()]

        if missing:
            st.warning(f"⚠️ Sheet berikut tidak ditemukan di file: {', '.join(missing)}")

        selected = st.sidebar.selectbox("Pilih sheet untuk ditampilkan", available)
        df = all_sheets[selected]

        st.subheader(f"📄 {selected}")
        st.write(f"Jumlah baris: {df.shape[0]} — Kolom: {df.shape[1]}")
        st.dataframe(df, use_container_width=True)

    except Exception as e:
        st.error(f"Error membaca Excel: {e}")
else:
    st.info("⬆️ Silakan upload file Excel kamu terlebih dahulu.")
