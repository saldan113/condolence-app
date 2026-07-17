import os
import smtplib
from datetime import datetime
from email.encoders import encode_base64
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from PIL import Image as PILImage  
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer

# ==========================================
# EMAIL CONFIGURATION (FULLY OPERATIONAL)
# ==========================================
SENDER_EMAIL = "condolences@ahmadiyya.us"  
SENDER_PASSWORD = "soqs qzbp brpp pdgm"  
CC_EMAIL = "national.ua@ahmadiyya.us"
EXCEL_FILENAME = "jamaat_directory.xlsx"  
# ==========================================

def get_honorific(gender_input):
    """Determines if the honorific should be Sahib or Sahiba."""
    clean_input = gender_input.strip().lower()
    if clean_input in ['m', 'male', 'sahib']:
        return "Sahib"
    elif clean_input in ['f', 'female', 'sahiba']:
        return "Sahiba"
    else:
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
        df = pd.read_excel(EXCEL_FILENAME)
        df['Jamaat'] = df['Jamaat'].astype(str).str.strip().str.lower()
        target_jamaat = jamaat_name.strip().lower()
        
        row = df[df['Jamaat'] == target_jamaat]
        if not row.empty:
            pres_email = str(row.iloc[0]['President Email']).strip()
            sec_email = str(row.iloc[0]['General Secretary Email']).strip()
            return [pres_email, sec_email]
        else:
            print(f"[Warning] Jamaat '{jamaat_name}' not found in Excel directory.")
            return []
    except Exception as e:
        print(f"[Error] Could not read Excel directory: {e}")
        return []

def send_condolence_email(recipient_emails, attachment_path, family_member, fam_honorific, relationship):
    if not recipient_emails:
        print("Skipping email broadcast: No local recipients found in Excel.")
        return

    recipients = [email for email in recipient_emails if email and email.lower() != 'nan']
    if not recipients:
        print("Skipping email broadcast: Recipient list is empty.")
        return

    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = ", ".join(recipients)
    msg['Cc'] = CC_EMAIL
    msg['Subject'] = f"Condolences - {relationship} of {family_member} {fam_honorific}"

    email_body = f"""Assalamo Alaikum,

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
        print(f"Failed to attach PDF: {e}")
        return

    all_recipients = recipients + [CC_EMAIL]

    try:
        print("Connecting to secure email server...")
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, all_recipients, msg.as_string())
        server.quit()
        print(f"Email successfully dispatched to local officers and CC'd to {CC_EMAIL}")
    except Exception as e:
        print(f"[SMTP Error] Failed to execute transmission: {e}")

def create_condolence_letter():
    print("--- Enhanced Condolence Letter Generator ---")
    # 1. Jamaat Name
    jamaat_name = input("Enter Jamaat Name: ").strip()
    
    # 2. Family Member details
    family_member = input("Enter the name of the family member of the deceased: ").strip()
    fam_gender = input("Is the family member Male (M) or Female (F)? ").strip()
    
    # 3. Deceased person's details
    deceased_member = input("Enter Deceased Person's Name: ").strip()
    dec_gender = input("Is the Deceased Person Male (M) or Female (F)? ").strip()
    
    # 4. Relationship to family member
    relationship = input("Enter relationship of the deceased to the family member (e.g. Mother, Father, Husband): ").strip()
    
    fam_honorific = get_honorific(fam_gender)
    dec_honorific = get_honorific(dec_gender)
    todays_date = datetime.now().strftime("%B %d, %Y")
    
    target_directory = os.path.join("Condolence letters", jamaat_name)
    os.makedirs(target_directory, exist_ok=True)
    pdf_filename = os.path.join(target_directory, f"{deceased_member}.pdf")
    
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
    
    try:
        story.append(Image("header.png", width=504, height=60))
        story.append(Spacer(1, 15))
    except Exception:
        print("[Warning] 'header.png' not found.")
    
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
    
    try:
        story.append(Spacer(1, 4))
        
        with PILImage.open("istirja.png") as img:
            orig_w, orig_h = img.size
        
        target_width = 150 
        aspect_ratio = orig_h / orig_w
        target_height = target_width * aspect_ratio
        
        story.append(Image("istirja.png", width=target_width, height=target_height, hAlign='CENTER'))
        story.append(Spacer(1, 4))
    except Exception as e:
        print(f"[Warning] Could not process 'istirja.png': {e}")
        story.append(Spacer(1, 4))
    
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
    
    try:
        story.append(Image("signature.png", width=120, height=45, hAlign='LEFT'))
        story.append(Spacer(1, 4))
    except Exception:
        pass
    
    signatory_text = (
        "Bilal Rana<br/>"
        "National Secretary Umur e Amma<br/>"
        "Jamaat Ahmadiyya USA<br/>"
        "National.ua@ahmadiyya.ua"
    )
    story.append(Paragraph(signatory_text, body_style))

    doc.build(story, onFirstPage=draw_footer, onLaterPages=draw_footer)
    print(f"\nSuccess! PDF saved locally into folder layout as: '{pdf_filename}'")
    
    # =========================================================================
    # LIVE SYSTEM ENGAGED: Emails will now broadcast immediately upon execution
    # =========================================================================
    emails = lookup_jamaat_emails(jamaat_name)
    send_condolence_email(emails, pdf_filename, family_member, fam_honorific, relationship)

if __name__ == "__main__":
    create_condolence_letter()
