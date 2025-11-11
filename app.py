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

        # Bersihkan spasi di semua kolom
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].astype(str).str.strip()

        st.subheader(f"📄 {selected}")

        # --- SEGMENTASI CUSTOMER PER KOTA ---
        if selected.lower().strip() == "segmentasi customer per kota":
            st.markdown("### 🏙️ Distribusi Pelanggan per Kota")
            city_col = df.columns[0]
            count_col = df.columns[1]
            fig = px.bar(
                df.sort_values(by=count_col, ascending=True),
                x=count_col,
                y=city_col,
                orientation="h",
                color=count_col,
                color_continuous_scale="blues",
                title="Jumlah Pelanggan Berdasarkan Kota",
            )
            st.plotly_chart(fig, use_container_width=True)

        # --- SEGMENTASI CUSTOMER PER PROVINSI ---
        elif selected.lower().strip() == "segmentasi customer per provins":
            st.markdown("### 🗺️ Distribusi Pelanggan per Provinsi")
            prov_col = df.columns[0]
            count_col = df.columns[1]
            fig = px.bar(
                df.sort_values(by=count_col, ascending=True),
                x=count_col,
                y=prov_col,
                orientation="h",
                color=count_col,
                color_continuous_scale="oranges",
                title="Jumlah Pelanggan Berdasarkan Provinsi",
            )
            st.plotly_chart(fig, use_container_width=True)

        # --- SHEET MARKETING ADS ---
        elif selected.lower().strip() == "marketing_ads":
            st.markdown("### 📈 Distribusi Biaya Iklan")
            numeric_cols = df.select_dtypes(include=["number"]).columns
            if len(numeric_cols) >= 1:
                fig = px.pie(df, names=df.columns[0], values=numeric_cols[0],
                             title="Pie Chart Marketing Ads")
                st.plotly_chart(fig, use_container_width=True)

        # --- SHEET PERTUMBUHAN PELANGGAN ---
        elif selected.lower().strip() == "pertumbuhan pelanggan":
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

        # --- SHEET CUSTOMER REGION (FIXED & FILTER READY) ---
        elif "customer region" in selected.lower():
            st.markdown("### 🌍 Data Pelanggan per Region")

            expected_cols = ["Nama Customer", "Kota", "Provinsi", "Terakhir Transaksi", "Jenis Produk yang Dibeli"]
            missing_cols = [c for c in expected_cols if c not in df.columns]
            if missing_cols:
                st.error(f"❌ Kolom berikut tidak ditemukan di sheet: {', '.join(missing_cols)}")
            else:
                # --- Filter Section di tampilan utama ---
                st.markdown("#### 🔍 Filter Data Pelanggan")

                col1, col2, col3 = st.columns(3)
                with col1:
                    kota_filter = st.multiselect("Pilih Kota", options=sorted(df["Kota"].unique()))
                with col2:
                    prov_filter = st.multiselect("Pilih Provinsi", options=sorted(df["Provinsi"].unique()))
                with col3:
                    produk_filter = st.multiselect("Pilih Jenis Produk", options=sorted(df["Jenis Produk yang Dibeli"].unique()))

                filtered_df = df.copy()
                if kota_filter:
                    filtered_df = filtered_df[filtered_df["Kota"].isin(kota_filter)]
                if prov_filter:
                    filtered_df = filtered_df[filtered_df["Provinsi"].isin(prov_filter)]
                if produk_filter:
                    filtered_df = filtered_df[filtered_df["Jenis Produk yang Dibeli"].isin(produk_filter)]

                # --- Tampilkan hasil filter ---
                st.markdown(f"Menampilkan **{len(filtered_df)}** data hasil filter:")
                st.dataframe(filtered_df, use_container_width=True)

        # --- DEFAULT ---
        else:
            st.dataframe(df, use_container_width=True)

    except Exception as e:
        st.error(f"⚠️ Error baca Excel: {e}")
