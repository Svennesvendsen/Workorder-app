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

            merged = workorders.merge(emails, on="WorkshopName", how="left")
            summary = merged.groupby(["WorkshopName"]).size().reset_index(name="OpenWorkorders")

            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                merged.to_excel(writer, sheet_name="DetaljeretData", index=False)
                summary.to_excel(writer, sheet_name="Oversigt", index=False)

                workbook = writer.book
                worksheet = workbook.add_worksheet("PrVærksted")
                writer.sheets["PrVærksted"] = worksheet

                row = 0
                for workshop, group in merged.groupby("WorkshopName"):
                    email = group['Email'].iloc[0]
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
                label="📥 Download rapport",
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

        st.sidebar.header("🔍 Vælg værksted")
        selected_workshop = st.sidebar.selectbox("Værksted", options=df["WorkshopName"].unique())

        filtered_df = df[df["WorkshopName"] == selected_workshop]

        st.subheader(f"📊 Dashboard for {selected_workshop}")
        st.metric("📦 Antal åbne workorders", len(filtered_df))

        if "CreationDate" in filtered_df.columns and "RepairDate" in filtered_df.columns:
            try:
                filtered_df["CreationDate"] = pd.to_datetime(filtered_df["CreationDate"])
                filtered_df["RepairDate"] = pd.to_datetime(filtered_df["RepairDate"])
                filtered_df["Behandlingstid (dage)"] = (filtered_df["RepairDate"] - filtered_df["CreationDate"]).dt.days
                avg_days = round(filtered_df["Behandlingstid (dage)"].mean(), 1)
                st.metric("⏱️ Gennemsnitlig behandlingstid", f"{avg_days} dage")
            except:
                st.warning("⚠️ Kunne ikke beregne behandlingstid.")

        st.subheader("📋 Detaljer")
        st.dataframe(filtered_df, use_container_width=True)

    else:
        st.info("Upload filer under 'Rapportgenerator' først.")
