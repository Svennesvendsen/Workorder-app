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

tab1, tab2 = st.tabs(["📊 Dashboard", "📄 PDF-rapport"])

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

    date_today = datetime.today().strftime("%d. %B %Y")
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"<b>Rapport genereret:</b> {date_today}", styles['Normal']))
    elements.append(Paragraph(f"<b>Kommentar:</b> {comment}", styles['Normal']))
    elements.append(Spacer(1, 24))

    elements.append(Paragraph(f"<b>Værksted:</b> {workshop_name}", styles['Title']))
    elements.append(Paragraph(f"<b>E-mail:</b> {email}", styles['Normal']))
    elements.append(Spacer(1, 12))

    display_cols = ['WONumber', 'AssetRegNo', 'CreationDate', 'RepairDate']
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
    st.markdown("### Upload workorders og e-mails")

    wo_file = st.file_uploader("📄 Upload workorders (Excel)", type=["xlsx"], key="wo")
    email_file = st.file_uploader("📧 Upload e-mails (Excel)", type=["xlsx"], key="email")

    if wo_file and email_file:
        try:
            wo_df = pd.read_excel(wo_file)
            email_df = pd.read_excel(email_file)

            required_cols = ['WONumber', 'WorkshopName', 'AssetRegNo', 'CreationDate', 'RepairDate']
            missing = [col for col in required_cols if col not in wo_df.columns]
            if missing:
                st.error(f"❌ Følgende kolonner mangler i workorder-filen: {', '.join(missing)}")
            else:
                merged_df = wo_df.merge(email_df, on="WorkshopName", how="left")
                merged_df["RepairDate"] = pd.to_datetime(merged_df["RepairDate"], errors="coerce")
                merged_df.loc[merged_df["RepairDate"].dt.year == 1900, "RepairDate"] = pd.NaT
                st.session_state["merged"] = merged_df

                st.sidebar.header("🎛 Visning")
                all_view = st.sidebar.checkbox("Vis alle værksteder samlet", value=True)

                if all_view:
                    st.subheader("📊 Samlet overblik")
                    st.metric("📦 Antal ordrer", len(merged_df))
                    st.metric("🏭 Antal værksteder", merged_df["WorkshopName"].nunique())

                    st.subheader("📈 Ordrer pr. værksted")
                    fig, ax = plt.subplots(figsize=(8, 4))
                    merged_df["WorkshopName"].value_counts().plot(kind="bar", ax=ax)
                    ax.set_ylabel("Ordrer")
                    st.pyplot(fig)

                    st.subheader("📋 Alle ordrer")
                    merged_df["CreationDate"] = pd.to_datetime(merged_df["CreationDate"], errors="coerce")
                    merged_df = merged_df.sort_values(by="CreationDate", ascending=True)
                    st.dataframe(merged_df, use_container_width=True)
                else:
                    selected_ws = st.sidebar.selectbox("Vælg værksted", options=merged_df["WorkshopName"].unique())
                    ws_df = merged_df[merged_df["WorkshopName"] == selected_ws]
                    st.subheader(f"📍 {selected_ws}")
                    st.metric("📦 Ordrer", len(ws_df))
                    if "CreationDate" in ws_df and "RepairDate" in ws_df:
                        try:
                            ws_df["CreationDate"] = pd.to_datetime(ws_df["CreationDate"])
                            ws_df["RepairDate"] = pd.to_datetime(ws_df["RepairDate"])
                            ws_df["Days"] = (ws_df["RepairDate"] - ws_df["CreationDate"]).dt.days
                            st.metric("⏱️ Gennemsnitlig behandlingstid", f"{round(ws_df['Days'].mean(), 1)} dage")
                        except:
                            pass
                    ws_df["CreationDate"] = pd.to_datetime(ws_df["CreationDate"], errors="coerce")
                    ws_df = ws_df.sort_values(by="CreationDate", ascending=True)
                    ws_df.loc[ws_df["RepairDate"].dt.year == 1900, "RepairDate"] = pd.NaT
                    st.dataframe(ws_df, use_container_width=True)
        except Exception as e:
            st.error(f"Fejl ved indlæsning: {e}")

with tab2:
    st.markdown("### Generér PDF for ét værksted")
    comment = st.text_input("🗒️ Kommentar til rapport", value="Vedhæftet ugentlig rapport")

    if "merged" in st.session_state:
        df = st.session_state["merged"]
        selected_ws = st.selectbox("Vælg værksted til PDF", options=df["WorkshopName"].unique(), key="pdf_ws")
        ws_df = df[df["WorkshopName"] == selected_ws]
        email = ws_df["Email"].iloc[0]
        ws_df["RepairDate"] = pd.to_datetime(ws_df["RepairDate"], errors="coerce")
        ws_df.loc[ws_df["RepairDate"].dt.year == 1900, "RepairDate"] = pd.NaT

        pdf_file = generate_pdf(ws_df, selected_ws, email, comment)
        st.download_button("📄 Download PDF", data=pdf_file,
                           file_name=f"rapport_{selected_ws.replace(' ', '_')}.pdf",
                           mime="application/pdf")
