import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="📊 CRM Dashboard", layout="wide")
st.title("📊 CRM Textek Dashboard")

# === FUNGSI LOAD DATA ===
@st.cache_data
def load_data(file_path, sheet_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Bersihkan kolom numerik yang mungkin pakai format 'Rp' atau koma
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).replace('[^0-9.-]', '', regex=True)
            try:
                df[col] = pd.to_numeric(df[col], errors='ignore')
            except:
                pass

        return df
    except Exception as e:
        st.error(f"Gagal load data dari sheet '{sheet_name}': {e}")
        return pd.DataFrame()

# === LOAD SEMUA SHEET ===
file_path = "CRM Salman Alfarizi.xlsx"

sheet1 = load_data(file_path, "Pelanggan Baru")
sheet2 = load_data(file_path, "VIP Buyer")

# === CEK JIKA DATA ADA ===
if sheet1.empty:
    st.warning("⚠️ Data sheet 'Pelanggan Baru' kosong atau tidak ditemukan.")
else:
    st.subheader("📈 Pertumbuhan Pelanggan Baru per Bulan")
    fig1 = px.line(sheet1, x='Bulan', y='Jumlah Pelanggan Didapat',
                   markers=True, title="Jumlah Pelanggan Didapat per Bulan")
    st.plotly_chart(fig1, use_container_width=True)

if sheet2.empty:
    st.warning("⚠️ Data sheet 'VIP Buyer' kosong atau tidak ditemukan.")
else:
    st.subheader("💎 Total Transaksi VIP Buyer")
    fig2 = px.bar(sheet2, x='Nama Buyer', y='Jumlah Transaksi',
                  title="Jumlah Transaksi per Buyer (VIP)")
    st.plotly_chart(fig2, use_container_width=True)

st.success("✅ Dashboard berhasil dimuat tanpa error!")
