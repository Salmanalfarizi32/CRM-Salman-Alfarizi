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
        selected = st.sidebar.selectbox("📑 Pilih Sheet", list(all_sheets.keys()))
        df = all_sheets[selected]

        # Bersihin data
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].astype(str).str.strip()

        # Format kolom uang otomatis
        money_cols = [c for c in df.columns if any(k in c.lower() for k in ["transaksi", "penjualan", "harga", "omzet"])]
        for col in money_cols:
            if pd.api.types.is_numeric_dtype(df[col]):
                continue
            df[col] = pd.to_numeric(df[col], errors="coerce")
            df[col] = df[col].apply(lambda x: f"Rp {x:,.0f}".replace(",", ".") if pd.notnull(x) else "-")

        st.subheader(f"📄 {selected}")
        st.dataframe(df, use_container_width=True)

        # --- PIE CHART UNTUK MARKETING ADS ---
        if selected.lower() == "marketing_ads":
            st.markdown("### 📈 Distribusi Biaya Iklan")
            numeric_cols = df.select_dtypes(include=["number"]).columns
            if len(numeric_cols) >= 1:
                fig = px.pie(df, names=df.columns[0], values=numeric_cols[0], title="Pie Chart Marketing Ads")
                st.plotly_chart(fig, use_container_width=True)

        # --- FIX FORMAT BULAN UNTUK PERTUMBUHAN PELANGGAN ---
        elif selected.lower() == "pertumbuhan pelanggan":
            if "Bulan" in df.columns:
                def smart_format(val):
                    try:
                        date_val = pd.to_datetime(val, errors="coerce")
                        if pd.notna(date_val):
                            return date_val.strftime("%B %Y")
                        return str(val)
                    except:
                        return str(val)
                df["Bulan"] = df["Bulan"].apply(smart_format)

            st.markdown("### 📊 Pertumbuhan Pelanggan per Bulan")
            st.dataframe(df, use_container_width=True)

            # auto chart kalau ada kolom jumlah
            jumlah_col = [c for c in df.columns if "jumlah" in c.lower()]
            if jumlah_col:
                fig = px.line(df, x="Bulan", y=jumlah_col[0], markers=True, title="Tren Pertumbuhan Pelanggan")
                st.plotly_chart(fig, use_container_width=True)

    except Exception as e:
        st.error(f"⚠️ Error baca Excel: {e}")
