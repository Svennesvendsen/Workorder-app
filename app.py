import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Workorder Rapport", layout="centered")
st.title("🔧 Workorder Rapportgenerator")

st.markdown("""
Upload to filer:
1. Excel med **aktive workorders**
2. Excel med **værksted-email mapping** (kolonner: `WorkshopName`, `Email`)

Appen genererer en rapport med:
- Detaljeret liste over alle åbne workorders med e-mails
- Oversigt over antal åbne ordrer pr. værksted
- ✅ Opdelt visning pr. værksted med én headerlinje
""")

workorder_file = st.file_uploader("📄 Upload workorder Excel-fil", type=["xlsx"])
email_file = st.file_uploader("📧 Upload værksted-email Excel-fil", type=["xlsx"])

if workorder_file and email_file:
    try:
        workorders = pd.read_excel(workorder_file)
        emails = pd.read_excel(email_file)

        # Merge data
        merged = workorders.merge(emails, on="WorkshopName", how="left")
        summary = merged.groupby(["WorkshopName", "Email"]).size().reset_index(name="OpenWorkorders")

        # Generer rapport
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            merged.to_excel(writer, sheet_name="DetaljeretData", index=False)
            summary.to_excel(writer, sheet_name="Oversigt", index=False)

            # Fanen: PrVærksted
            workbook = writer.book
            worksheet = workbook.add_worksheet("PrVærksted")
            writer.sheets["PrVærksted"] = worksheet

            row = 0
            for (workshop, email), group in merged.groupby(["WorkshopName", "Email"]):
                # Overskrift: værkstedsnavn og e-mail
                worksheet.write(row, 0, f"🏭 {workshop} – {email}")
                row += 2
                # Kolonneoverskrifter
                for col_num, col_name in enumerate(group.columns):
                    worksheet.write(row, col_num, col_name)
                row += 1
                # Rækker
                for _, data_row in group.iterrows():
                    for col_num, value in enumerate(data_row):
                        worksheet.write(row, col_num, str(value))
                    row += 1
                row += 2  # Luft før næste værksted

        output.seek(0)

        st.success("✅ Rapport genereret!")
        st.download_button(
            label="📥 Download rapport med værkstedsvisning",
            data=output,
            file_name="rapport_med_pr_vaerksted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Noget gik galt: {e}")
