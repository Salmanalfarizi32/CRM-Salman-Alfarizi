import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="📊 CRM Dashboard", layout="wide")
st.title("📊 CRM Textek Dashboard")

# === Upload file Excel ===
uploaded_file = st.sidebar.file_uploader("📂 Upload file Excel kamu (.xlsx)", type=["xlsx"])

if uploaded_file:
    # Baca semua sheet
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    selected_sheet = st.sidebar.selectbox("Pilih sheet", sheet_names)

    df = pd.read_excel(xls, sheet_name=selected_sheet)

    # Bersihkan data dari simbol Rp, koma, spasi
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).replace('[^0-9.,-]', '', regex=True)
            df[col] = df[col].replace('', None)
            try:
                df[col] = pd.to_numeric(df[col], errors='ignore')
            except:
                pass

    st.subheader(f"📄 Data dari Sheet: {selected_sheet}")
    st.dataframe(df, use_container_width=True)

    # === Visualisasi otomatis ===
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    if len(numeric_cols) >= 1:
        st.subheader("📊 Visualisasi Otomatis")
        x_axis = st.selectbox("Pilih kolom X", df.columns)
        y_axis = st.selectbox("Pilih kolom Y (numerik)", numeric_cols)

        if y_axis:
            fig = px.bar(df, x=x_axis, y=y_axis, title=f"{y_axis} berdasarkan {x_axis}")
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Tidak ada kolom numerik untuk divisualisasikan.")

else:
    st.warning("⬅️ Upload file Excel kamu dulu di sidebar.")
