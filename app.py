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
    for _, row in df.iterrows():
        link = row.get("Platform link", "")
        wo = str(row["WONumber"])
        if pd.notna(link) and link:
            wo_link = f'<link href="{link}">{wo}</link>'
        else:
            wo_link = wo
        table_data.append([wo_link, str(row["AssetRegNo"]), str(row["CreationDate"]), str(row["RepairDate"])])

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

# (rest of the app code continues unchanged...)
