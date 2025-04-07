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

try:
    import win32com.client as win32
except ImportError:
    win32 = None

st.set_page_config(page_title="Workorder App", layout="wide")
st.title("🔧 Workorder App")

tab1, tab2 = st.tabs(["📄 Rapportgenerator", "📊 Dashboard"])

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

def send_email_with_pdf(recipient, pdf_bytes, workshop_name):
    if win32 is None:
        st.error("win32com er ikke tilgængelig. Kan ikke sende e-mail fra denne maskine.")
        return

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = f"Workorder rapport – {workshop_name}"
    mail.Body = f"Hej\n\nVedhæftet ugentlig rapport for åbne workorders hos {workshop_name}.\n\nMvh\nAutomatisk system"

    temp_path = os.path.join(os.environ["TEMP"], "rapport.pdf")
    with open(temp_path, "wb") as f:
        f.write(pdf_bytes.read())
    mail.Attachments.Add(temp_path)
    mail.Send()
    os.remove(temp_path)

with tab1:
    st.info("Brug dashboard-fanen for PDF og mailfunktion.")

with tab2:
    st.markdown("### 📊 Dashboard + PDF/E-mail funktion")

    workorder_file = st.file_uploader("📄 Upload workorder Excel-fil", type=["xlsx"], key="wo_file2")
    email_file = st.file_uploader("📧 Upload værksted-email Excel-fil", type=["xlsx"], key="email_file2")
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

        if st.button("✉️ Send rapport til værkstedets e-mail"):
            send_email_with_pdf(email, pdf_data, selected_ws)
            st.success(f"📬 Rapport sendt til {email}")
