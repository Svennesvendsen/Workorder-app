import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

st.set_page_config(page_title="Workorder App", layout="wide")
st.title("🔧 Workorder App")

# Faner
tab1, tab2 = st.tabs(["📄 Rapportgenerator", "📊 Dashboard"])

with tab1:
    st.markdown("""
    Upload to filer:
    1. Excel med **aktive workorders**
    2. Excel med **værksted-email mapping** (kolonner: `WorkshopName`, `Email`)
    """)

    workorder_file = st.file_uploader("📄 Upload workorder Excel-fil", type=["xlsx"], key="wo_file")
    email_file = st.file_uploader("📧 Upload værksted-email Excel-fil", type=["xlsx"], key="email_file")

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
                    worksheet.write(row, 0, f"🏭 {workshop} – {email}")
                    row += 2
                    for col_num, col_name in enumerate(group.columns):
                        worksheet.write(row, col_num, col_name)
                    row += 1
                    for _, data_row in group.iterrows():
                        for col_num, value in enumerate(data_row):
                            worksheet.write(row, col_num, str(value))
                        row += 1
                    row += 2

            output.seek(0)

            st.success("✅ Rapport genereret!")
            st.download_button(
                label="📥 Download rapport med værkstedsvisning",
                data=output,
                file_name="rapport_med_pr_vaerksted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.session_state["dashboard_data"] = merged

        except Exception as e:
            st.error(f"Noget gik galt: {e}")

with tab2:
    if "dashboard_data" in st.session_state:
        df = st.session_state["dashboard_data"]

        with st.sidebar:
            st.header("🔍 Filtrér data")
            workshops = st.multiselect("Vælg værksted(er)", options=df['WorkshopName'].unique(), default=df['WorkshopName'].unique())
            asset_filter = st.text_input("Søg efter AssetRegNo")

        filtered_df = df[df['WorkshopName'].isin(workshops)]
        if asset_filter:
            filtered_df = filtered_df[filtered_df['AssetRegNo'].astype(str).str.contains(asset_filter, case=False)]

        st.metric("📦 Antal åbne workorders", len(filtered_df))
        st.metric("🏭 Antal værksteder", filtered_df['WorkshopName'].nunique())

        st.subheader("📈 Åbne ordrer pr. værksted")
        count_by_ws = filtered_df['WorkshopName'].value_counts()
        fig, ax = plt.subplots()
        count_by_ws.plot(kind='bar', ax=ax)
        ax.set_ylabel("Antal ordrer")
        ax.set_xlabel("Værksted")
        ax.set_title("Åbne ordrer pr. værksted")
        st.pyplot(fig)

        st.subheader("📋 Detaljeret workorder-liste")
        st.dataframe(filtered_df, use_container_width=True)
    else:
        st.info("Upload filer under 'Rapportgenerator' først.")

