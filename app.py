import time
import os
import base64
import json
import gspread
from google.oauth2.service_account import Credentials
from flask import Flask, request, jsonify, render_template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import threading
import smtplib
from dotenv import load_dotenv
from email.mime.application import MIMEApplication
import pdfkit
import tempfile


load_dotenv()
SHEET_ID = os.environ.get("SHEET_ID")
ENCODED_CREDS = os.environ.get("GOOGLE_CREDENTIALS_JSON_BASE64")
CHECK_INTERVAL = 5  

SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
SMTP_USERNAME = os.environ.get("EMAIL_ADDRESS")
SMTP_PASSWORD = os.environ.get("EMAIL_PASSWORD")
SENDER_EMAIL = SMTP_USERNAME

app = Flask(__name__)

processed_users = set()  
last_processed_row = 1 
actively_processing = False  

# Column indexes (0-based)
COL_TIMESTAMP = 0
COL_EMAIL = 1
COL_FIRST_NAME = 2
COL_LAST_NAME = 3
COL_COURSE = 4
COL_PERIOD = 5
COL_MOBILE = 6
COL_STATUS = 7  # Status will be written to column H
COL_PROCESSED_TIME = 8  # Processing timestamp will be written to column I

def get_credentials():
    if not ENCODED_CREDS:
        raise ValueError("GOOGLE_CREDENTIALS_JSON_BASE64 environment variable is not set")
    
    creds_json = base64.b64decode(ENCODED_CREDS).decode('utf-8')
    creds_info = json.loads(creds_json)
    
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    return creds

def send_certificate_email(recipient_email, name, course):
    try:
        msg = MIMEMultipart('alternative')
        msg['From'] = SENDER_EMAIL
        msg['To'] = recipient_email
        msg['Subject'] = 'Certificate of Completion'

        text = f"Dear {name},\n\nPlease find your certificate details below:\n\nCompleted Course: {course}"
        html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Your Certificate</title>
            <style>
                body {{
                    font-family: 'Arial', sans-serif;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    min-height: 100vh;
                    background-color: #f0f0f0;
                    margin: 0;
                }}

                .certificate-container {{
                    background-color: #fff;
                    border: 10px solid #4CAF50;
                    padding: 40px;
                    box-shadow: 0 0 20px rgba(0, 0, 0, 0.1);
                    width: 800px;
                    max-width: 95%;
                    text-align: center;
                }}

                .certificate-title {{
                    font-size: 2.5em;
                    color: #333;
                    margin-bottom: 20px;
                }}

                .presented-to {{
                    font-size: 1.2em;
                    margin-bottom: 10px;
                }}

                .recipient-name {{
                    font-size: 2em;
                    font-weight: bold;
                    color: #007bff;
                    margin-bottom: 15px;
                }}

                .for-completion {{
                    font-size: 1.2em;
                    margin-bottom: 10px;
                }}

                .course-name {{
                    font-size: 1.6em;
                    font-style: italic;
                    color: #555;
                    margin-bottom: 25px;
                }}

                .signature-area {{
                    display: flex;
                    justify-content: space-around;
                    margin-top: 30px;
                }}

                .signature {{
                    border-top: 2px dashed #ccc;
                    padding-top: 10px;
                    width: 250px;
                }}

                .date {{
                    font-size: 0.9em;
                    color: #777;
                    margin-top: 10px;
                }}
            </style>
        </head>
        <body>
            <div class="certificate-container">
                <h2 class="certificate-title">Certificate of Completion</h2>
                <p class="presented-to">This certificate is presented to</p>
                <h1 class="recipient-name">{name}</h1>
                <p class="for-completion">for successfully completing the</p>
                <p class="course-name">{course}</p>
                <div class="signature-area">
                    <div class="signature">
                        (Signature)
                        <p class="date">Date: {datetime.now().strftime('%B %d, %Y')}</p>
                    </div>
                    <div class="signature">
                        (Instructor/Organizer)
                        <p class="date"></p>
                    </div>
                </div>
            </div>
        </body>
        </html>
        """

        part1 = MIMEText(text, 'plain')
        part2 = MIMEText(html, 'html')

        msg.attach(part1)
        msg.attach(part2)
        BASE_DIR = os.path.abspath(os.path.dirname(__file__))
        path_to_wkhtmltopdf = os.path.join(BASE_DIR, 'wkhtmltopdf.exe')

        config = pdfkit.configuration(wkhtmltopdf=path_to_wkhtmltopdf)


        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
            pdfkit.from_string(html, tmp_pdf.name, configuration=config)
            tmp_pdf_path = tmp_pdf.name

        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = recipient_email
        msg['Subject'] = 'Your Certificate of Completion'

        body = f"Dear {name},\n\nCongratulations Please find your certificate of course completation of '{course}' it has been attached as a PDF.\n\nBest regards,\n Mayur Dhoble"
        msg.attach(MIMEText(body, 'plain'))

        with open(tmp_pdf_path, "rb") as f:
            pdf_attachment = MIMEApplication(f.read(), _subtype="pdf")
            pdf_attachment.add_header('Content-Disposition', 'attachment', filename=f"{name}_Certificate.pdf")
            msg.attach(pdf_attachment)

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.sendmail(SENDER_EMAIL, recipient_email, msg.as_string())

        return True

    except Exception as e:
        print(f"Error sending email: {e}")
        return False

def update_status(sheet, row_index, status):
    try:
        # Update status in column H (index 7)
        sheet.update_cell(row_index, COL_STATUS + 1, status)
        
        # Update timestamp in column I (index 8)
        sheet.update_cell(row_index, COL_PROCESSED_TIME + 1, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        
        time.sleep(1)  # Avoid API rate limiting
    except Exception as e:
        print(f"Error updating status: {e}")

def generate_user_key(email, first_name, last_name):
    """Create a unique key for each user to prevent duplicates"""
    email = email.strip().lower()
    first_name = first_name.strip().lower()
    last_name = last_name.strip().lower()
    return f"{email}:{first_name}:{last_name}"

def initialize_processed_users(sheet):
    global processed_users
    try:
        all_rows = sheet.get_all_values()
        
        for row in all_rows[1:]:  # Skip header row
            if len(row) > COL_LAST_NAME:  # Make sure row has enough columns
                email = row[COL_EMAIL].strip()
                first_name = row[COL_FIRST_NAME].strip()
                last_name = row[COL_LAST_NAME].strip() if len(row) > COL_LAST_NAME else ""
                
                if email and first_name:
                    user_key = generate_user_key(email, first_name, last_name)
                    processed_users.add(user_key)
        
        print(f"Initialized with {len(processed_users)} processed users")
    except Exception as e:
        print(f"Error initializing processed users: {e}")

def is_already_processed(row):
    if len(row) <= COL_LAST_NAME:
        return False
    
    email = row[COL_EMAIL].strip()
    first_name = row[COL_FIRST_NAME].strip()
    last_name = row[COL_LAST_NAME].strip() if len(row) > COL_LAST_NAME else ""
    
    user_key = generate_user_key(email, first_name, last_name)
    return user_key in processed_users

def monitor_spreadsheet():
    global last_processed_row, processed_users, actively_processing
    print("Starting spreadsheet monitoring...")
    
    try:
        creds = get_credentials()
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SHEET_ID).sheet1
        
        initialize_processed_users(sheet)
        all_rows = sheet.get_all_values()
        last_processed_row = len(all_rows)
        print(f"Starting with {last_processed_row} existing rows")
        
    except Exception as e:
        print(f"Error during initial setup: {e}")
    
    while True:
        if actively_processing:
            time.sleep(1)
            continue
            
        try:
            actively_processing = True
            
            creds = get_credentials()
            client = gspread.authorize(creds)
            sheet = client.open_by_key(SHEET_ID).sheet1
            
            all_rows = sheet.get_all_values()
            current_row_count = len(all_rows)
            
            if current_row_count > last_processed_row:
                print(f"New submission detected: Row {current_row_count}")
                
                row = all_rows[current_row_count-1]
                
                if len(row) > COL_MOBILE:  # Make sure row has enough columns
                    email = row[COL_EMAIL].strip()
                    first_name = row[COL_FIRST_NAME].strip()
                    last_name = row[COL_LAST_NAME].strip()  
                    full_name = f"{first_name} {last_name}"
                    course = row[COL_COURSE].strip()          
                    course_period = row[COL_PERIOD].strip() if len(row) > COL_PERIOD else ""
                    mobile = row[COL_MOBILE].strip() if len(row) > COL_MOBILE else ""
                    
                    if not is_already_processed(row):
                        print(f"Processing new submission: {full_name}, {email}, {course}")
                        
                        update_status(sheet, current_row_count, "Processing")
                        
                        if send_certificate_email(email, full_name, course):
                            update_status(sheet, current_row_count, "Certificate Sent")
                            print(f"Certificate sent successfully to {email}")
                        else:
                            update_status(sheet, current_row_count, "Email Failed")
                            print(f"Failed to send certificate to {email}")
                        
                        # Add to processed set after successful processing
                        user_key = generate_user_key(email, first_name, last_name)
                        processed_users.add(user_key)
                    else:
                        print(f"Skipping already processed: {full_name}")
                        update_status(sheet, current_row_count, "Duplicate - No Certificate Sent")
                
                last_processed_row = current_row_count
            
        except Exception as e:
            print(f"Error monitoring spreadsheet: {e}")
        finally:
            actively_processing = False
        
        time.sleep(CHECK_INTERVAL)

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/admin')
def admin():
    return render_template('admin.html')

@app.route('/status')
def status():
    try:
        creds = get_credentials()
        client = gspread.authorize(creds)
        
        sheet = client.open_by_key(SHEET_ID).sheet1
        all_rows = sheet.get_all_values()
        
        # Check if header row exists and adjust accordingly
        has_header = len(all_rows) > 0 and all_rows[0][0].lower() == "timestamp"
        start_index = 1 if has_header else 0
        
        recent_submissions = []
        for row in all_rows[start_index:]:
            if len(row) > COL_COURSE:  # Make sure row has enough columns for basic info
                email = row[COL_EMAIL].strip() if len(row) > COL_EMAIL else ""
                first_name = row[COL_FIRST_NAME].strip() if len(row) > COL_FIRST_NAME else ""
                last_name = row[COL_LAST_NAME].strip() if len(row) > COL_LAST_NAME else ""
                full_name = f"{first_name} {last_name}"
                
                submission = {
                    'timestamp': row[COL_TIMESTAMP] if len(row) > COL_TIMESTAMP else "",
                    'email': email,
                    'name': full_name,
                    'course': row[COL_COURSE] if len(row) > COL_COURSE else "",
                    'period': row[COL_PERIOD] if len(row) > COL_PERIOD else "",
                    'mobile': row[COL_MOBILE] if len(row) > COL_MOBILE else "",
                    'status': row[COL_STATUS] if len(row) > COL_STATUS else 'Pending',
                    'processed_at': row[COL_PROCESSED_TIME] if len(row) > COL_PROCESSED_TIME else ''
                }
                recent_submissions.append(submission)
        
        # Show the most recent 10 submissions
        recent_submissions = recent_submissions[-10:]
        
        return render_template('status.html', submissions=recent_submissions)
    except Exception as e:
        return f"Error retrieving status: {str(e)}"

@app.route('/health')
def health():
    return jsonify({
        "status": "up", 
        "timestamp": datetime.now().isoformat(),
        "processed_users_count": len(processed_users),
        "last_processed_row": last_processed_row,
        "actively_processing": actively_processing
    })

@app.route('/processed_users')
def view_processed_users():
    return jsonify({
        "processed_users": list(processed_users),
        "count": len(processed_users)
    })

if __name__ == '__main__':
    required_vars = ["SHEET_ID", "GOOGLE_CREDENTIALS_JSON_BASE64", "EMAIL_ADDRESS", "EMAIL_PASSWORD"]
    missing_vars = [var for var in required_vars if not os.environ.get(var)]
    
    if missing_vars:
        print(f"Error: Missing required environment variables: {', '.join(missing_vars)}")
    else:
        monitor_thread = threading.Thread(target=monitor_spreadsheet, daemon=True)
        monitor_thread.start()
        
        port = int(os.environ.get("PORT", 5000))
        app.run(host='0.0.0.0', port=port, debug=False)
