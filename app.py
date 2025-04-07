import streamlit as st
import pandas as pd
from io import BytesIO
import smtplib
from email.message import EmailMessage

st.set_page_config(page_title="Workorder Rapport", layout="centered")
st.title("🔧 Ugentlig Workorder Rapport + Mail")

st.markdown("""
Upload to filer:
1. Excel med **aktive workorders**
2. Excel med **værksted-email mapping** (WorkshopName + Email)

Efter rapporten er genereret, kan du sende den direkte via mail 📤
""")

workorder_file = st.file_uploader("Upload workorder Excel-fil", type=["xlsx"])
email_file = st.file_uploader("Upload værksted-email Excel-fil", type=["xlsx"])

if 'report_bytes' not in st.session_state:
    st.session_state['report_bytes'] = None

if workorder_file and email_file:
    try:
        workorders = pd.read_excel(workorder_file)
        emails = pd.read_excel(email_file)

        merged = workorders.merge(emails, on="WorkshopName", how="left")
        summary = merged.groupby(["WorkshopName", "Email"]).size().reset_index(name="OpenWorkorders")

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            merged.to_excel(writer, sheet_name='DetaljeretData', index=False)
            summary.to_excel(writer, sheet_name='Oversigt', index=False)
        output.seek(0)

        st.success("✅ Rapport genereret!")
        st.session_state['report_bytes'] = output.read()
        st.download_button(
            label="📥 Download rapport",
            data=st.session_state['report_bytes'],
            file_name="rapport_med_emails.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Fejl under behandling: {e}")

# Send mail sektion
if st.session_state['report_bytes']:
    with st.expander("✉️ Send rapport via mail"):
        smtp_server = st.text_input("SMTP-server", value="smtp.office365.com")
        smtp_port = st.number_input("SMTP-port", value=587)
        sender_email = st.text_input("Afsender email", value="villads@pnorental.dk")
        password = st.text_input("Adgangskode / App-adgangskode", type="password")
        to_emails = st.text_input("Modtagere (kommasepareret)", value="villads@pnorental.dk, peter@pnorental.dk")
        subject = st.text_input("Emne", value="Månedlig Workorder Rapport")
        body = st.text_area("Besked", value="Hej\n\nHermed månedlig rapport over åbne workorders pr. værksted.\n\nMvh\nAutomatisk system")

        if st.button("📤 Send rapport via mail"):
            try:
                msg = EmailMessage()
                msg['Subject'] = subject
                msg['From'] = sender_email
                msg['To'] = [email.strip() for email in to_emails.split(",")]
                msg.set_content(body)

                msg.add_attachment(st.session_state['report_bytes'],
                                   maintype='application',
                                   subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                   filename="rapport_med_emails.xlsx")

                with smtplib.SMTP(smtp_server, smtp_port) as smtp:
                    smtp.starttls()
                    smtp.login(sender_email, password)
                    smtp.send_message(msg)

                st.success("📬 Mail sendt!")
            except Exception as e:
                st.error(f"Fejl ved afsendelse: {e}")
