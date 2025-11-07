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
    return pd.read_excel(path, sheet_name=None)

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

        # --- FORMAT KHUSUS ---
        money_keywords = ["transaksi", "penjualan", "harga", "omzet"]
        for col in df.columns:
            if any(key in col.lower() for key in money_keywords):
                df[col] = pd.to_numeric(df[col], errors="coerce")
                df[col] = df[col].apply(
                    lambda x: f"Rp {x:,.0f}".replace(",", ".") if pd.notnull(x) else "-"
                )

        st.subheader(f"📄 {selected}")

        # --- KHUSUS UNTUK SHEET MARKETING ADS ---
        if selected == "Marketing_ads":
            st.write("📈 Distribusi Biaya Iklan per Platform")
            numeric_cols = df.select_dtypes(include=["int64", "float64"]).columns
            if len(numeric_cols) >= 1:
                value_col = numeric_cols[0]
                fig = px.pie(df, names=df.columns[0], values=value_col, title="Pie Chart Marketing Ads")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Tidak ada kolom numerik untuk dibuat chart.")
            st.dataframe(df, use_container_width=True)

        # --- KHUSUS UNTUK SHEET PERTUMBUHAN PELANGGAN ---
        elif selected == "Pertumbuhan Pelanggan":
            if "Bulan" in df.columns:
                try:
                    df["Bulan"] = pd.to_datetime(df["Bulan"], errors="coerce")
                    if df["Bulan"].notna().any():
                        df["Bulan"] = df["Bulan"].dt.strftime("%B %Y")
                    else:
                        df["Bulan"] = df["Bulan"].astype(str)
                except:
                    df["Bulan"] = df["Bulan"].astype(str)

            st.write("📊 Pertumbuhan Pelanggan per Bulan")
            st.dataframe(df, use_container_width=True)

        # --- UNTUK SHEET LAIN ---
        else:
            st.dataframe(df, use_container_width=True)

    except Exception as e:
        st.error(f"Error baca Excel: {e}")
