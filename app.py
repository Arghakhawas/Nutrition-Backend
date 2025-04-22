from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd
import os
from datetime import datetime
from werkzeug.utils import secure_filename
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import traceback
from dotenv import load_dotenv
import time

# === Load .env Variables ===
load_dotenv()

# === Configuration ===
UPLOAD_FOLDER = "uploads"
LOG_FOLDER = "logs"
SENDER_EMAIL = "argha820@gmail.com"
SENDER_PASSWORD = os.environ.get("EMAIL_PASSWORD")

# === App Initialization ===
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
CORS(app)

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(LOG_FOLDER, exist_ok=True)

PERMANENT_SIGNATURE = """
---
Regards,  
Argha Khawas  
Nutrition Expert  
📞 Contact: 9073357827
"""

# === Email Sender ===
def send_email(to_email, subject, body, image=None):
    try:
        msg = MIMEMultipart('mixed')
        msg['From'] = SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = subject
        msg['Reply-To'] = SENDER_EMAIL
        msg['Return-Path'] = SENDER_EMAIL
        msg['X-Priority'] = '3'
        msg['X-Mailer'] = 'Python Flask'

        # Plain text (without signature to avoid duplication)
        msg.attach(MIMEText(body, 'plain'))

        # Attachment
        if image:
            img_data = image.read()
            part = MIMEApplication(img_data, Name=image.filename)
            part['Content-Disposition'] = f'attachment; filename="{image.filename}"'
            msg.attach(part)

        # HTML version with signature
        html_body_content = body.replace("\n", "<br>")
        html_signature = PERMANENT_SIGNATURE.replace("\n", "<br>")
        html_body = f"""
        <html>
            <body>
                <p>{html_body_content}</p>
                <p>{html_signature}</p>
            </body>
        </html>
        """
        msg.attach(MIMEText(html_body, 'html'))

        # SMTP sending
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, to_email, msg.as_string())
        server.quit()

    except Exception as e:
        print(f"❌ Error sending email to {to_email}: {e}")
        raise e

# === Send Endpoint ===
@app.route('/send', methods=['POST'])
def send_messages():
    try:
        excel = request.files.get('excel')
        if not excel or 'message' not in request.form:
            return jsonify({"status": "❌ Excel file and message are required."}), 400

        image = request.files.get('image')
        message_template = request.form['message']
        filename = secure_filename(excel.filename)
        excel_path = os.path.join(UPLOAD_FOLDER, filename)
        excel.save(excel_path)

        df = pd.read_excel(excel_path)

        if 'Email' not in df.columns or 'Name' not in df.columns:
            return jsonify({"status": "❌ Excel must contain 'Email' and 'Name' columns."}), 400

        log_filename = f"log_{datetime.now().strftime('%Y%m%d%H%M%S')}.txt"
        log_path = os.path.join(LOG_FOLDER, log_filename)

        with open(log_path, "w", encoding="utf-8") as log:
            for _, row in df.iterrows():
                try:
                    name = str(row['Name']).strip()
                    email = str(row['Email']).strip()
                    if pd.isna(name) or pd.isna(email):
                        continue
                    personalized_msg = message_template.replace("(Name)", name)
                    send_email(email, "Message from Argha", personalized_msg, image)
                    log.write(f"✅ Sent to {email}\n")
                    time.sleep(2)  # Delay to avoid spam flagging
                except Exception as e:
                    log.write(f"❌ Failed to {row.get('Email')}: {e}\n")
                    continue

        return jsonify({
            "status": "✅ Emails sent successfully.",
            "log_url": f"/logs/{log_filename}"
        }), 200

    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": f"❌ Internal Server Error: {e}"}), 500

# === Serve Logs ===
@app.route('/logs/<path:filename>')
def download_log(filename):
    return send_from_directory(LOG_FOLDER, filename)

# === Health Check ===
@app.route('/')
def home():
    return jsonify({"message": "✅ Nutrition Backend is live and running!"})

# === Main Runner ===
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
