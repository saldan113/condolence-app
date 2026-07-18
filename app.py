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

# ==========================================
# EMAIL CONFIGURATION
# ==========================================
SENDER_EMAIL = "condolences@ahmadiyya.us"  
SENDER_PASSWORD = "soqs qzbp brpp pdgm"  
CC_EMAIL = "national.ua@ahmadiyya.us"
EXCEL_FILENAME = "jamaat_directory.xlsx"  
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
            st.error(f"Excel file '{EXCEL_FILENAME}' not found in the repository.")
            return []
        df = pd.read_excel(EXCEL_FILENAME)
        df['Jamaat'] = df['Jamaat'].astype(str).str.strip().str.lower()
        target_jamaat = jamaat_name.strip().lower()
        
        row = df[df['Jamaat'] == target_jamaat]
        if not row.empty:
            pres_email = str(row.iloc[0]['President Email']).strip()
            sec_email = str(row.iloc[0]['General Secretary Email']).strip()
            return [pres_email, sec_email]
        else:
            st.warning(f"Jamaat '{jamaat_name}' not found in Excel directory.")
            return []
    except Exception as e:
        st.error(f"Could not read Excel directory: {e}")
        return []

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
            part.add_header(
                "Content-Disposition",
                f"attachment; filename={os.path.basename(attachment_path)}",
            )
            msg.attach(part)
    except Exception as e:
        st.error(f"Failed to attach PDF to email: {e}")
        return False

    all_recipients = recipients + [CC_EMAIL]

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, all_recipients, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"SMTP Server Error: {e}")
        return False

# ==========================================
# STREAMLIT WEB INTERFACE
# ==========================================
st.set_page_config(page_title="Condolence Letter Generator", layout="centered")
st.title("Umur-e-Amma Condolence Letter Portal")
st.write("Fill out the fields below to generate the official PDF and email it directly to the local Sadr and General Secretary.")

# Load the Jamaat names list from the Excel spreadsheet for the dropdown
jamaat_options = []
if os.path.exists(EXCEL_FILENAME):
    try:
        df_list = pd.read_excel(EXCEL_FILENAME)
        if 'Jamaat' in df_list.columns:
            # Extract unique values, drop blanks, and sort alphabetically
            jamaat_options = df_list['Jamaat'].dropna().astype(str).str.strip().unique().tolist()
            jamaat_options.sort()
        else:
            st.error("Error: Could not find a 'Jamaat' column heading in your Excel spreadsheet.")
    except Exception as e:
        st.error(f"Error loading dropdown menu options from Excel: {e}")
else:
    st.error(f"Error: '{EXCEL_FILENAME}' is missing from the app environment.")

# ==========================================
# APP INPUT FLOW
# ==========================================
# 1. Dynamic Dropdown for Jamaat Name selection
if jamaat_options:
    jamaat_name = st.selectbox("Select Jamaat Name", jamaat_options)
else:
    # Safe text input fallback just in case the file reading has a hitch
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
        st.error("Please fill in all the text fields before generating.")
    else:
        fam_honorific = get_honorific(fam_gender)
        dec_honorific = get_honorific(dec_gender)
        todays_date = datetime.now().strftime("%B %d, %Y")
        
        pdf_filename = f"Condolence Letter - {deceased_member} - {todays_date}.pdf"
        
        # Build Document Layout
        doc = SimpleDocTemplate(
            pdf_filename,
            pagesize=letter,
            rightMargin=54,
            leftMargin=54,
            topMargin=40,
            bottomMargin=90
        )
        
        styles = getSampleStyleSheet()
        body_style = ParagraphStyle(
            'LetterBody', parent=styles['Normal'], fontName='Helvetica', fontSize=11, leading=18, spaceAfter=8
        )
        translation_style = ParagraphStyle(
            'TranslationStyle', parent=styles['Normal'], fontName='Helvetica-Oblique', fontSize=9.5, leading=15, alignment=1, spaceAfter=10
        )
        
        story = []
        
        # Header Graphic
        if os.path.exists("header.png"):
            story.append(Image("header.png", width=504, height=60))
            story.append(Spacer(1, 15))
        
        story.append(Paragraph(f"<b>Date:</b> {todays_date}", body_style))
        story.append(Spacer(1, 8))
        
        story.append(Paragraph(f"<b>Dear Respected {family_member} {fam_honorific},</b>", body_style))
        story.append(Paragraph("<i>Assalamu Alaikum wa Rahmatullahi wa Barakatuhu,</i>", body_style))
        story.append(Spacer(1, 4))
        
        body_text_1 = (
            f"On behalf of Jamaat Ahmadiyya USA, we express our deepest condolences on the "
            f"passing of your beloved family member, Respected <b>{deceased_member} {dec_honorific}</b>."
        )
        story.append(Paragraph(body_text_1, body_style))
        
        # Arabic Prayer Image Handling
        if os.path.exists("istirja.png"):
            try:
                story.append(Spacer(1, 4))
                with PILImage.open("istirja.png") as img:
                    orig_w, orig_h = img.size
                target_width = 150 
                aspect_ratio = orig_h / orig_w
                target_height = target_width * aspect_ratio
                story.append(Image("istirja.png", width=target_width, height=target_height, hAlign='CENTER'))
                story.append(Spacer(1, 4))
            except Exception:
                story.append(Spacer(1, 4))
        
        # Centered Translation Text
        translation_text = "“Surely, to Allah we belong, and to Him we return,” (The Holy Qur'an, 2:157)."
        story.append(Paragraph(translation_text, translation_style))
        
        body_text_2 = (
            "May Allah the Almighty grant them <i>Maghfirat</i> (forgiveness) and <i>Janat al-Firdous</i> "
            "(loftiest paradise). May He grant the bereaved solace and patience."
        )
        story.append(Paragraph(body_text_2, body_style))
        story.append(Spacer(1, 8))
        
        story.append(Paragraph("With heartfelt prayers,", body_style))
        story.append(Paragraph("Wasalaam", body_style))
        story.append(Spacer(1, 4))
        
        # Signature Image
        if os.path.exists("signature.png"):
            story.append(Image("signature.png", width=120, height=45, hAlign='LEFT'))
            story.append(Spacer(1, 4))
        
        signatory_text = (
            "Bilal Rana<br/>"
            "National Secretary Umur e Amma<br/>"
            "Jamaat Ahmadiyya USA<br/>"
            "National.ua@ahmadiyya.ua"
        )
        story.append(Paragraph(signatory_text, body_style))

        # Build Document
        doc.build(story, onFirstPage=draw_footer, onLaterPages=draw_footer)
        
        # Lookup and Transmit
        with st.spinner("Looking up local leadership contacts and broadcasting email..."):
            emails = lookup_jamaat_emails(jamaat_name)
            if emails:
                success = send_condolence_email(emails, pdf_filename, family_member, fam_honorific, relationship)
                if success:
                    st.success(f"Letter successfully generated and emailed to {', '.join(emails)}!")
                    
                    with open(pdf_filename, "rb") as file:
                        st.download_button(
                            label="Download generated PDF copy",
                            data=file,
                            file_name=pdf_filename,
                            mime="application/pdf"
                        )
            else:
                st.error("Could not send email because the Jamaat name was not found in the Excel directory.")
        
        # Clean up local generated file
        if os.path.exists(pdf_filename):
            os.remove(pdf_filename)
