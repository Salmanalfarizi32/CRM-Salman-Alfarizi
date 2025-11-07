import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

st.set_page_config(page_title="📊 CRM Dashboard", layout="wide")
st.title("📊 CRM Textek Dashboard")

EXCEL_PATH = Path("data/CRM Salman Alfarizi.xlsx")

if not EXCEL_PATH.exists():
    st.error(f"❌ File tidak ditemukan: {EXCEL_PATH}")
else:
    try:
        all_sheets = pd.read_excel(EXCEL_PATH, sheet_name=None)
        selected = st.sidebar.selectbox("📑 Pilih Sheet", list(all_sheets.keys()))
        df = all_sheets[selected].copy()

        # Bersihkan data string
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].astype(str).str.strip()

        # --- FORMAT KHUSUS UNTUK NOMINAL ---
        money_keywords = ["transaksi", "penjualan", "harga", "omzet"]
        for col in df.columns:
            if any(k in col.lower() for k in money_keywords):
                df[col] = pd.to_numeric(df[col], errors="coerce")
                df[col] = df[col].apply(
                    lambda x: f"Rp {x:,.0f}".replace(",", ".") if pd.notnull(x) else "-"
                )

        st.subheader(f"📄 {selected}")

        # --- SHEET MARKETING ADS ---
        if selected.lower() == "marketing_ads":
            st.markdown("### 📈 Distribusi Biaya Iklan")
            numeric_cols = df.select_dtypes(include=["number"]).columns
            if len(numeric_cols) >= 1:
                fig = px.pie(df, names=df.columns[0], values=numeric_cols[0],
                             title="Pie Chart Marketing Ads")
                st.plotly_chart(fig, use_container_width=True)

        # --- SHEET PERTUMBUHAN PELANGGAN ---
        elif selected.lower() == "pertumbuhan pelanggan":
            if "Bulan" in df.columns:
                def smart_format(val):
                    try:
                        d = pd.to_datetime(val, errors="coerce")
                        if pd.notna(d):
                            return d.strftime("%B %Y")
                        return str(val)
                    except:
                        return str(val)
                df["Bulan"] = df["Bulan"].apply(smart_format)

            jumlah_col = [c for c in df.columns if "jumlah" in c.lower()]
            if jumlah_col:
                fig = px.line(df, x="Bulan", y=jumlah_col[0], markers=True,
                              title="📊 Tren Pertumbuhan Pelanggan")
                st.plotly_chart(fig, use_container_width=True)

        # --- TAMPILKAN DATAFRAME CUMA SEKALI ---
        st.dataframe(df, use_container_width=True)

    except Exception as e:
        st.error(f"⚠️ Error baca Excel: {e}")
