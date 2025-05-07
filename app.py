import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
import os

st.set_page_config(page_title="Workorder Dashboard", layout="wide")
st.title("🔧 Workorder Dashboard + PDF")

tab1, tab2 = st.tabs(["📊 Dashboard", "📄 PDF Report"])

def generate_pdf(df, workshop_name, email, comment):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()

    try:
        logo_path = os.path.join(os.path.dirname(__file__), "PNO_logo_2018_RGB.png")
        if os.path.exists(logo_path):
            logo = RLImage(logo_path, width=120, height=120)
            logo.hAlign = 'CENTER'
            elements.append(logo)
    except:
        pass

    date_today = datetime.today().strftime("%d %B %Y")
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"<b>Report generated:</b> {date_today}", styles['Normal']))
    elements.append(Paragraph(f"<b>Note:</b> {comment}", styles['Normal']))
    elements.append(Spacer(1, 24))

    elements.append(Paragraph(f"<b>Workshop:</b> {workshop_name}", styles['Title']))
    elements.append(Paragraph(f"<b>Email:</b> {email}", styles['Normal']))
    elements.append(Spacer(1, 12))

    display_cols = ['WONumber', 'AssetRegNo', 'CreationDate', 'RepairDate']
    table_data = [display_cols]
    style = styles["Normal"]
    for _, row in df.iterrows():
        link = row.get("Platform link", "")
        wo = str(row["WONumber"])
        if pd.notna(link) and link:
            wo_link = Paragraph(f'<link href="{link}">{wo}</link>', style)
        else:
            wo_link = Paragraph(wo, style)
        table_data.append([
            wo_link,
            Paragraph(str(row["AssetRegNo"]), style),
            Paragraph(str(row["CreationDate"]), style),
            Paragraph(str(row["RepairDate"]), style)
        ])

    table = Table(table_data, hAlign='LEFT')
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
    ]))
    elements.append(table)
    doc.build(elements)
    buffer.seek(0)
    return buffer

with tab1:
    st.markdown("### Upload workorders and workshop emails")

    wo_file = st.file_uploader("📄 Upload workorders (Excel)", type=["xlsx"], key="wo")
    email_file = st.file_uploader("📧 Upload workshop emails (Excel)", type=["xlsx"], key="email")

    if wo_file and email_file:
        try:
            wo_df = pd.read_excel(wo_file)
            email_df = pd.read_excel(email_file)

            required_cols = ['WONumber', 'WorkshopName', 'AssetRegNo', 'CreationDate', 'RepairDate']
            missing = [col for col in required_cols if col not in wo_df.columns]
            if missing:
                st.error(f"❌ Missing columns in workorder file: {', '.join(missing)}")
            else:
                merged_df = wo_df.merge(email_df, on="WorkshopName", how="left")
                merged_df["RepairDate"] = pd.to_datetime(merged_df["RepairDate"], errors="coerce")
                merged_df.loc[merged_df["RepairDate"].dt.year == 1900, "RepairDate"] = pd.NaT
                st.session_state["merged"] = merged_df

                st.sidebar.header("🎛 View options")
                all_view = st.sidebar.checkbox("Show all workshops combined", value=True)

                if all_view:
                    st.subheader("📊 Overall Overview")
                    st.metric("📦 Number of workorders", len(merged_df))
                    st.metric("🏭 Number of workshops", merged_df["WorkshopName"].nunique())

                    st.subheader("📈 Workorders per workshop")
                    fig, ax = plt.subplots(figsize=(8, 4))
                    merged_df["WorkshopName"].value_counts().plot(kind="bar", ax=ax)
                    ax.set_ylabel("Workorders")
                    st.pyplot(fig)

                    st.subheader("📋 All workorders")
                    merged_df["CreationDate"] = pd.to_datetime(merged_df["CreationDate"], errors="coerce")
                    merged_df = merged_df.sort_values(by="CreationDate", ascending=True)
                    st.dataframe(merged_df, use_container_width=True)
                else:
                    selected_ws = st.sidebar.selectbox("Select workshop", options=merged_df["WorkshopName"].unique())
                    ws_df = merged_df[merged_df["WorkshopName"] == selected_ws]
                    st.subheader(f"📍 {selected_ws}")
                    st.metric("📦 Workorders", len(ws_df))
                    if "CreationDate" in ws_df and "RepairDate" in ws_df:
                        try:
                            ws_df["CreationDate"] = pd.to_datetime(ws_df["CreationDate"])
                            ws_df["RepairDate"] = pd.to_datetime(ws_df["RepairDate"])
                            ws_df["Days"] = (ws_df["RepairDate"] - ws_df["CreationDate"]).dt.days
                            st.metric("⏱️ Avg. repair time", f"{round(ws_df['Days'].mean(), 1)} days")
                        except:
                            pass
                    ws_df["CreationDate"] = pd.to_datetime(ws_df["CreationDate"], errors="coerce")
                    ws_df = ws_df.sort_values(by="CreationDate", ascending=True)
                    ws_df.loc[ws_df["RepairDate"].dt.year == 1900, "RepairDate"] = pd.NaT
                    st.dataframe(ws_df, use_container_width=True)
        except Exception as e:
            st.error(f"Error loading files: {e}")

with tab2:
    st.markdown("### Generate PDF for a single workshop")
    comment = st.text_input("🗒️ Note for the report", value="Open workorders")

    if "merged" in st.session_state:
        df = st.session_state["merged"]
        selected_ws = st.selectbox("Select workshop for PDF", options=df["WorkshopName"].unique(), key="pdf_ws")
        ws_df = df[df["WorkshopName"] == selected_ws]
        email = ws_df["Email"].iloc[0]
        ws_df["RepairDate"] = pd.to_datetime(ws_df["RepairDate"], errors="coerce")
        ws_df.loc[ws_df["RepairDate"].dt.year == 1900, "RepairDate"] = pd.NaT

        pdf_file = generate_pdf(ws_df, selected_ws, email, comment)
        st.download_button("📄 Download PDF", data=pdf_file,
                           file_name=f"report_{selected_ws.replace(' ', '_')}.pdf",
                           mime="application/pdf")
