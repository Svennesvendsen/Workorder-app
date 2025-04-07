import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="Workorder Dashboard", layout="wide")
st.title("📊 Workorder Dashboard")

st.markdown("""
Upload din Excel-fil med aktive workorders. Dashboardet viser:
- Antal åbne ordrer pr. værksted
- Visualiseringer og interaktiv tabel
""")

uploaded_file = st.file_uploader("📄 Upload workorder Excel-fil", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        # UI: filtre
        with st.sidebar:
            st.header("🔍 Filtrér data")
            workshops = st.multiselect("Vælg værksted(er)", options=df['WorkshopName'].unique(), default=df['WorkshopName'].unique())
            asset_filter = st.text_input("Søg efter AssetRegNo")
        
        # Filtrér data
        filtered_df = df[df['WorkshopName'].isin(workshops)]
        if asset_filter:
            filtered_df = filtered_df[filtered_df['AssetRegNo'].astype(str).str.contains(asset_filter, case=False)]

        # KPI'er
        total_orders = len(filtered_df)
        unique_workshops = filtered_df['WorkshopName'].nunique()
        st.metric("📦 Antal åbne workorders", total_orders)
        st.metric("🏭 Antal værksteder", unique_workshops)

        # Plot
        st.subheader("📈 Åbne ordrer pr. værksted")
        count_by_ws = filtered_df['WorkshopName'].value_counts()
        fig, ax = plt.subplots()
        count_by_ws.plot(kind='bar', ax=ax)
        ax.set_ylabel("Antal ordrer")
        ax.set_xlabel("Værksted")
        ax.set_title("Åbne ordrer pr. værksted")
        st.pyplot(fig)

        # Interaktiv tabel
        st.subheader("📋 Detaljeret workorder-liste")
        st.dataframe(filtered_df, use_container_width=True)
    except Exception as e:
        st.error(f"Noget gik galt ved indlæsning af filen: {e}")
