import os
import smtplib
from datetime import datetime
from email.encoders import encode_base64
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from PIL import Image as PILImage  
import streamlit as st
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer
import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# CONFIGURATION
# ==========================================
SENDER_EMAIL = "condolences@ahmadiyya.us"  
SENDER_PASSWORD = "soqs qzbp brpp pdgm"  
CC_EMAIL = "national.ua@ahmadiyya.us"
EXCEL_FILENAME = "jamaat_directory.xlsx"  
GOOGLE_SHEET_NAME = "Condolence_Letters_Log"
CREDS_FILE = "google_creds.json"
# ==========================================

def get_honorific(gender_input):
    if gender_input == 'Male':
        return "Sahib"
    elif gender_input == 'Female':
        return "Sahiba"
    return "Sahib/Sahiba"

def draw_footer(canvas, doc):
    canvas.saveState()
    try:
        canvas.drawImage("footer.png", 54, 40, width=504, height=40)
    except Exception:
        pass
    canvas.restoreState()

def lookup_jamaat_emails(jamaat_name):
    try:
        if not os.path.exists(EXCEL_FILENAME):
            st.error(f"Excel file '{EXCEL_FILENAME}' not found.")
            return []
        df = pd.read_excel(EXCEL_FILENAME)
        df['Jamaat'] = df['Jamaat'].astype(str).str.strip().str.lower()
        target_jamaat = jamaat_name.strip().lower()
        row = df[df['Jamaat'] == target_jamaat]
        if not row.empty:
            return [str(row.iloc[0]['President Email']).strip(), str(row.iloc[0]['General Secretary Email']).strip()]
        return []
    except Exception as e:
        st.error(f"Error reading local registry: {e}")
        return []

def log_to_google_sheets(jamaat, deceased, family_member, relationship):
    """Appends a tracking log line including the relationship layout to the remote cloud Google Sheet."""
    try:
        if not os.path.exists(CREDS_FILE):
            st.error("Google authentication file missing from app node.")
            return
        
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_file(CREDS_FILE, scopes=scopes)
        client = gspread.authorize(creds)
        
        sheet = client.open(GOOGLE_SHEET_NAME).sheet1
        today_str = datetime.now().strftime("%Y-%m-%d")
        # Appends Date, Jamaat, Deceased, Family Member, and Relationship
        sheet.append_row([today_str, jamaat, deceased, family_member, relationship])
    except Exception as e:
        st.error(f"Failed to log entry to cloud database: {e}")

def get_sheet_data():
    """Reads all rows from the Google Sheet for reporting analytics."""
    try:
        if not os.path.exists(CREDS_FILE):
            return pd.DataFrame()
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_file(CREDS_FILE, scopes=scopes)
        client = gspread.authorize(creds)
        sheet = client.open(GOOGLE_SHEET_NAME).sheet1
        records = sheet.get_all_records()
        return pd.DataFrame(records)
    except Exception:
        return pd.DataFrame()

def send_condolence_email(recipient_emails, attachment_path, family_member, fam_honorific, relationship):
    recipients = [email for email in recipient_emails if email and str(email).lower() != 'nan']
    if not recipients:
        st.error("Email broadcast failed: Recipient list is empty or invalid.")
        return False

    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = ", ".join(recipients)
    msg['Cc'] = CC_EMAIL
    msg['Subject'] = f"Condolences - {relationship} of {family_member} {fam_honorific}"

    email_body = """Assalamo Alaikum,

Dear Respected Sadr sahib and Secretary sahib,

I pray this message finds you all in the best of health and high spirits.

Attached is the condolences letter prepared on behalf of the ‘Umūr-e-‘Amma Department regarding the recent passing. Kindly review the letter and ensure it is shared promptly with the relevant families and members within your local Jamaat, in accordance with Jamaat protocols.

If you have any questions or require any further clarification, please feel free to reach out.

We humbly request your prayers for the departed soul and for the family during this difficult time.

Jazākumullāh,

Wassalām,

'Umūr-e-‘Amma Department"""

    msg.attach(MIMEText(email_body, 'plain', 'utf-8'))

    try:
        with open(attachment_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(attachment_path)}")
            msg.attach(part)
    except Exception as e:
        st.error(f"Failed to attach PDF: {e}")
        return False

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, recipients + [CC_EMAIL], msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"SMTP Server Error: {e}")
        return False

# ==========================================
# STREAMLIT INTERFACE WITH DUAL TABS
# ==========================================
st.set_page_config(page_title="Condolence Letter Portal", layout="centered")

# Navigation Tabs
tab1, tab2 = st.tabs(["📝 Generate Letter", "📊 Live Stats Dashboard"])

# Load Jamaat options
jamaat_options = []
if os.path.exists(EXCEL_FILENAME):
    try:
        df_list = pd.read_excel(EXCEL_FILENAME)
        if 'Jamaat' in df_list.columns:
            jamaat_options = df_list['Jamaat'].dropna().astype(str).str.strip().unique().tolist()
            jamaat_options.sort()
    except Exception:
        pass

with tab1:
    st.title("Letter Dispatch Portal")
    if jamaat_options:
        jamaat_name = st.selectbox("Select Jamaat Name", jamaat_options)
    else:
        jamaat_name = st.text_input("Enter Jamaat Name").strip()

    st.subheader("Family Member Details")
    family_member = st.text_input("Enter the name of the family member of the deceased").strip()
    fam_gender = st.selectbox("Is the family member Male or Female?", ["Male", "Female"])

    st.subheader("Deceased Details")
    deceased_member = st.text_input("Enter Deceased Person's Name").strip()
    dec_gender = st.selectbox("Is the Deceased Person Male or Female?", ["Male", "Female"])
    relationship = st.text_input("Enter relationship of the deceased to the family member (e.g. Mother, Father, Husband)").strip()

    if st.button("Generate & Email Condolence Letter", type="primary"):
        if not jamaat_name or not family_member or not deceased_member or not relationship:
            st.error("Please fill in all the fields before executing.")
        else:
            fam_honorific = get_honorific(fam_gender)
            dec_honorific = get_honorific(dec_gender)
            todays_date = datetime.now().strftime("%B %d, %Y")
            
            pdf_filename = f"Condolence Letter - {deceased_member} - {todays_date}.pdf"
            
            doc = SimpleDocTemplate(pdf_filename, pagesize=letter, rightMargin=54, leftMargin=54, topMargin=40, bottomMargin=90)
            styles = getSampleStyleSheet()
            body_style = ParagraphStyle('LetterBody', parent=styles['Normal'], fontName='Helvetica', fontSize=11, leading=18, spaceAfter=8)
            translation_style = ParagraphStyle('TranslationStyle', parent=styles['Normal'], fontName='Helvetica-Oblique', fontSize=9.5, leading=15, alignment=1, spaceAfter=10)
            
            story = []
            if os.path.exists("header.png"):
                story.append(Image("header.png", width=504, height=60))
                story.append(Spacer(1, 15))
            
            story.append(Paragraph(f"<b>Date:</b> {todays_date}", body_style))
            story.append(Spacer(1, 8))
            story.append(Paragraph(f"<b>Dear Respected {family_member} {fam_honorific},</b>", body_style))
            story.append(Paragraph("<i>Assalamu Alaikum wa Rahmatullahi wa Barakatuhu,</i>", body_style))
            story.append(Spacer(1, 4))
            
            body_text_1 = f"On behalf of Jamaat Ahmadiyya USA, we express our deepest condolences on the passing of your beloved family member, Respected <b>{deceased_member} {dec_honorific}</b>."
            story.append(Paragraph(body_text_1, body_style))
            
            if os.path.exists("istirja.png"):
                try:
                    story.append(Spacer(1, 4))
                    with PILImage.open("istirja.png") as img:
                        orig_w, orig_h = img.size
                    target_width = 150 
                    target_height = target_width * (orig_h / orig_w)
                    story.append(Image("istirja.png", width=target_width, height=target_height, hAlign='CENTER'))
                    story.append(Spacer(1, 4))
                except Exception:
                    pass
            
            translation_text = "“Surely, to Allah we belong, and to Him we return,” (The Holy Qur'an, 2:157)."
            story.append(Paragraph(translation_text, translation_style))
            
            body_text_2 = "May Allah the Almighty grant them <i>Maghfirat</i> (forgiveness) and <i>Janat al-Firdous</i> (loftiest paradise). May He grant the bereaved solace and patience."
            story.append(Paragraph(body_text_2, body_style))
            story.append(Spacer(1, 8))
            story.append(Paragraph("With heartfelt prayers,", body_style))
            story.append(Paragraph("Wasalaam", body_style))
            story.append(Spacer(1, 4))
            
            if os.path.exists("signature.png"):
                story.append(Image("signature.png", width=120, height=45, hAlign='LEFT'))
                story.append(Spacer(1, 4))
            
            signatory_text = "Bilal Rana<br/>National Secretary Umur e Amma<br/>Jamaat Ahmadiyya USA<br/>National.ua@ahmadiyya.ua"
            story.append(Paragraph(signatory_text, body_style))
            doc.build(story, onFirstPage=draw_footer, onLaterPages=draw_footer)
            
            with st.spinner("Processing cloud logging and broadcasting email..."):
                emails = lookup_jamaat_emails(jamaat_name)
                if emails:
                    if send_condolence_email(emails, pdf_filename, family_member, fam_honorific, relationship):
                        # Log to Google Sheets securely upon email success (now includes relationship variable)
                        log_to_google_sheets(jamaat_name, deceased_member, family_member, relationship)
                        st.success(f"Letter generated and transmitted to {', '.join(emails)}!")
                        
                        with open(pdf_filename, "rb") as file:
                            st.download_button(label="Download generated PDF copy", data=file, file_name=pdf_filename, mime="application/pdf")
                else:
                    st.error("Could not send email. Verify the selected Jamaat name exists in the Excel file registry.")
            
            if os.path.exists(pdf_filename):
                os.remove(pdf_filename)

with tab2:
    st.title("Departmental Analytics Dashboard")
    st.write("Live, on-demand metrics pulled securely from the master cloud log sheet.")
    
    df_metrics = get_sheet_data()
    if not df_metrics.empty:
        st.metric(label="Total Condolence Letters Dispatched Overall", value=len(df_metrics))
        
        st.subheader("Letters Distributed per Local Jamaat")
        jamaat_counts = df_metrics['Jamaat'].value_counts().reset_index()
        jamaat_counts.columns = ['Jamaat Name', 'Letters Issued']
        st.dataframe(jamaat_counts, use_container_width=True)
        
        st.subheader("Master Historical Audit Log")
        st.dataframe(df_metrics, use_container_width=True)
    else:
        st.info("No logs found. Once letters are processed, real-time aggregate stats will appear here.")
