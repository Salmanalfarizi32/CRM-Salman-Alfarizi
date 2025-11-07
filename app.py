import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

st.set_page_config(page_title="CRM Dashboard - Excel Auto Mode", layout="wide")
st.title("📊 CRM Dashboard (Auto Excel Mode)")

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
        selected = st.sidebar.selectbox("Pilih sheet", available)
        df = all_sheets[selected]

        st.subheader(f"📄 {selected}")

        # === 1️⃣ Sheet VIP BUYER ===
        if selected == "VIP BUYER":
            if "Total Transaksi" in df.columns:
                df["Total Transaksi"] = df["Total Transaksi"].apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))
            st.dataframe(df, use_container_width=True)

        # === 2️⃣ Sheet Marketing_ads ===
        elif selected == "Marketing_ads":
            st.dataframe(df, use_container_width=True)
            numeric_cols = df.select_dtypes("number").columns.tolist()
            if len(numeric_cols) >= 1:
                pie_col = st.selectbox("Pilih kolom untuk Pie Chart", df.columns)
                value_col = st.selectbox("Pilih nilai untuk Pie Chart", numeric_cols)
                fig = px.pie(df, names=pie_col, values=value_col, title="📊 Distribusi Marketing Ads")
                st.plotly_chart(fig, use_container_width=True)

        # === 3️⃣ Sheet Pertumbuhan Pelanggan ===
        elif selected == "Pertumbuhan Pelanggan":
            if "Bulan" in df.columns:
                df["Bulan"] = pd.to_datetime(df["Bulan"], errors="coerce")
                df["Bulan"] = df["Bulan"].dt.strftime("%B %Y")  # contoh: "Oktober 2025"
            st.dataframe(df, use_container_width=True)
            if df.select_dtypes("number").shape[1] > 0:
                y_col = st.selectbox("Pilih kolom angka untuk grafik", df.select_dtypes("number").columns)
                fig = px.line(df, x="Bulan", y=y_col, markers=True, title="📈 Pertumbuhan Pelanggan per Bulan")
                st.plotly_chart(fig, use_container_width=True)

        # === 4️⃣ Sheet lainnya (tampilkan tabel saja) ===
        else:
            st.dataframe(df, use_container_width=True)

    except Exception as e:
        st.error(f"Error baca Excel: {e}")
