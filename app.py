import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="Workorder App", layout="wide")
st.title("🔧 Workorder App – PDF-rapport")

tab1 = st.container()

def generate_pdf(df, workshop_name, email, comment):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()

    try:
        logo = RLImage("PNO_logo_2018_RGB.png", width=120, height=120)
        logo.hAlign = 'CENTER'
        elements.append(logo)
    except:
        pass

    date_today = datetime.today().strftime("%d. %B %Y")
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"<b>Rapport genereret:</b> {date_today}", styles['Normal']))
    elements.append(Paragraph(f"<b>Kommentar:</b> {comment}", styles['Normal']))
    elements.append(Spacer(1, 24))

    elements.append(Paragraph(f"<b>Værksted:</b> {workshop_name}", styles['Title']))
    elements.append(Paragraph(f"<b>E-mail:</b> {email}", styles['Normal']))
    elements.append(Spacer(1, 12))

    display_cols = ['WorkorderID', 'AssetRegNo', 'CreationDate', 'RepairDate', 'Amount']
    table_data = [display_cols] + df[display_cols].astype(str).values.tolist()

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
    st.markdown("### 📄 Generér PDF-rapport for værksted")

    workorder_file = st.file_uploader("📄 Upload workorder Excel-fil", type=["xlsx"], key="wo_file")
    email_file = st.file_uploader("📧 Upload værksted-email Excel-fil", type=["xlsx"], key="email_file")
    comment = st.text_input("🗒️ Kommentar til rapport", value="Vedhæftet ugentlig rapport")

    if workorder_file and email_file:
        try:
            df = pd.read_excel(workorder_file)
            emails = pd.read_excel(email_file)
            merged = df.merge(emails, on="WorkshopName", how="left")
            st.session_state["merged_data"] = merged
        except Exception as e:
            st.error(f"Fejl ved indlæsning: {e}")

    if "merged_data" in st.session_state:
        df = st.session_state["merged_data"]
        selected_ws = st.selectbox("Vælg værksted", options=df["WorkshopName"].unique())
        ws_df = df[df["WorkshopName"] == selected_ws]
        email = ws_df["Email"].iloc[0]

        st.metric("📦 Antal åbne workorders", len(ws_df))

        pdf_data = generate_pdf(ws_df, selected_ws, email, comment)
        st.download_button("📄 Download PDF-rapport", data=pdf_data,
                           file_name=f"rapport_{selected_ws.replace(' ', '_')}.pdf",
                           mime="application/pdf")
