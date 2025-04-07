
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Workorder Rapport", layout="centered")
st.title("🔧 Ugentlig Workorder Rapport")

st.markdown("Upload din Excel-fil med aktive workorders. Appen genererer en rapport med antal åbne workorders pr. værksted.")

uploaded_file = st.file_uploader("Upload Excel-fil", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df = xls.parse(xls.sheet_names[0])

        # Generer oversigt over antal åbne workorders pr. værksted
        workshop_counts = df['WorkshopName'].value_counts().reset_index()
        workshop_counts.columns = ['WorkshopName', 'OpenWorkorders']

        # Gem begge som ny Excel-fil i hukommelsen
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Workorders', index=False)
            workshop_counts.to_excel(writer, sheet_name='Oversigt', index=False)
        output.seek(0)

        st.success("Rapport genereret!")
        st.download_button(
            label="📎 Download rapport",
            data=output,
            file_name="rapport_over_abne_workorders.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Noget gik galt under behandlingen af filen: {e}")
