import os
from datetime import datetime
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# ==========================================
# REPORTING CONFIGURATION
# ==========================================
REPORT_RECIPIENT = "national.ua@ahmadiyya.us"
SENDER_EMAIL = "condolences@ahmadiyya.us"
SENDER_PASSWORD = "soqs qzbp brpp pdgm"
GOOGLE_SHEET_NAME = "Condolence_Letters_Log"
CREDS_FILE = "google_creds.json"
# ==========================================

def run_monthly_report():
    if not os.path.exists(CREDS_FILE):
        print("Missing credentials file.")
        return

    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=scopes)
    client = gspread.authorize(creds)
    sheet = client.open(GOOGLE_SHEET_NAME).sheet1
    
    df = pd.DataFrame(sheet.get_all_records())
    if df.empty:
        print("No letters logged yet.")
        return

    df['Date'] = pd.to_datetime(df['Date'])
    current_month = datetime.now().month
    current_year = datetime.now().year
    current_month_name = datetime.now().strftime("%B")

    monthly_data = df[(df['Date'].dt.month == current_month) & (df['Date'].dt.year == current_year)]
    total_letters = len(monthly_data)

    email_body = f"Assalamo Alaikum,\n\nHere is the Umur-e-Amma monthly condolence letter summary report for {current_month_name} {current_year}.\n\n"
    email_body += f"Total Letters Dispatched This Month: {total_letters}\n\n"
    
    if total_letters > 0:
        email_body += "Breakdown by Local Jamaat:\n"
        jamaat_counts = monthly_data['Jamaat'].value_counts()
        for jamaat, count in jamaat_counts.items():
            email_body += f" * {jamaat}: {count} letter(s)\n"
    else:
        email_body += "No condolence letters were generated during this monthly period.\n"

    email_body += "\nJazakumullah,\n\n'Umur-e-Amma Automated Reporting System"

    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = REPORT_RECIPIENT
    msg['Subject'] = f"Monthly Condolence Letters Summary - {current_month_name} {current_year}"
    msg.attach(MIMEText(email_body, 'plain', 'utf-8'))

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, [REPORT_RECIPIENT], msg.as_string())
        server.quit()
        print("Monthly statistics report delivered successfully.")
    except Exception as e:
        print(f"Failed to transmit email: {e}")

if __name__ == "__main__":
    run_monthly_report()
