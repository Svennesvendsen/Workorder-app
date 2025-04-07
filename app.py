import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Workorder Rapport", layout="centered")
st.title("🔧 Udvidet Workorder Rapport")

st.markdown("""
Upload to filer:
1. Excel med **aktive workorders**
2. Excel med **værksted-email mapping** (WorkshopName + Email)
""")

workorder_file = st.file_uploader("Upload workorder Excel-fil", type=["xlsx"])
email_file = st.file_uploader("Upload værksted-email Excel-fil", type=["xlsx"])

if workorder_file and email_file:
    try:
        # Læs data
        workorders = pd.read_excel(workorder_file)
        emails = pd.read_excel(email_file)

        # Merge email-adresser ind i workorders
        merged = workorders.merge(emails, on="WorkshopName", how="left")

        # Lav oversigt
        summary = merged.groupby(["WorkshopName", "Email"]).size().reset_index(name="OpenWorkorders")

        # Gem begge som ny Excel-fil
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            merged.to_excel(writer, sheet_name='DetaljeretData', index=False)
            summary.to_excel(writer, sheet_name='Oversigt', index=False)
        output.seek(0)

        st.success("Rapport genereret!")
        st.download_button(
            label="📎 Download rapport",
            data=output,
            file_name="rapport_med_emails.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Noget gik galt under behandlingen af filerne: {e}")
