import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

st.set_page_config(page_title="📊 CRM Dashboard", layout="wide")
st.title("📊 CRM Textek Dashboard")

# Path file Excel
EXCEL_PATH = Path("data/CRM Salman Alfarizi.xlsx")

# Baca file Excel
if not EXCEL_PATH.exists():
    st.error(f"❌ File tidak ditemukan: {EXCEL_PATH}")
else:
    try:
        all_sheets = pd.read_excel(EXCEL_PATH, sheet_name=None)

        st.sidebar.header("📑 Navigasi Sheet")
        selected = st.sidebar.selectbox("Pilih sheet", list(all_sheets.keys()))

        df = all_sheets[selected]

        # --- Bersihkan data agar rapi ---
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].astype(str).str.strip()

        # Format kolom uang otomatis
        money_cols = [c for c in df.columns if any(k in c.lower() for k in ["transaksi", "penjualan", "harga", "omzet"])]
        for col in money_cols:
            df[col] = pd.to_numeric(df[col], errors="coerce")
            df[col] = df[col].apply(lambda x: f"Rp {x:,.0f}".replace(",", ".") if pd.notnull(x) else "-")

        st.subheader(f"📄 {selected}")
        st.dataframe(df, use_container_width=True)

        # --- Visualisasi otomatis berdasarkan nama sheet ---
        if selected.lower() == "marketing_ads":
            st.markdown("### 📈 Distribusi Biaya Iklan")
            numeric_cols = df.select_dtypes(include=["number"]).columns
            if len(numeric_cols) >= 1:
                fig = px.pie(df, names=df.columns[0], values=numeric_cols[0], title="Pie Chart Marketing Ads")
                st.plotly_chart(fig, use_container_width=True)

        elif selected.lower() == "pertumbuhan pelanggan":
            if "Bulan" in df.columns:
                df["Bulan"] = df["Bulan"].astype(str)
                df["Bulan"] = df["Bulan"].replace("None", "")
                # ubah format ke 'Oktober 2025' dsb
                try:
                    df["Bulan"] = pd.to_datetime(df["Bulan"], errors="coerce").dt.strftime("%B %Y")
                except:
                    pass
            st.markdown("### 📊 Pertumbuhan Pelanggan per Bulan")
            fig = px.line(df, x="Bulan", y=df.columns[-1], markers=True)
            st.plotly_chart(fig, use_container_width=True)

    except Exception as e:
        st.error(f"⚠️ Error baca Excel: {e}")
