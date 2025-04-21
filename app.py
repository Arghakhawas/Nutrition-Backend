from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd
import os
import smtplib
import time
from email.message import EmailMessage
from werkzeug.utils import secure_filename

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = "uploads"
LOG_FOLDER = "logs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(LOG_FOLDER, exist_ok=True)

EMAIL_ADDRESS = "argha820@gmail.com"
EMAIL_PASSWORD = "your_app_password"  # Use app-specific password

@app.route("/send", methods=["POST"])
def send_emails():
    file = request.files.get("excel")
    image = request.files.get("image")
    message_template = request.form.get("message", "").strip()

    if not file or not message_template:
        return jsonify({"status": "Missing file or message"}), 400

    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    df = pd.read_excel(filepath)

    if "Email" not in df.columns:
        return jsonify({"status": "Missing 'Email' column in Excel"}), 400

    has_name = "Name" in df.columns
    attachment_path = None

    if image:
        attachment_filename = secure_filename(image.filename)
        attachment_path = os.path.join(UPLOAD_FOLDER, attachment_filename)
        image.save(attachment_path)

    log_file_path = os.path.join(LOG_FOLDER, f"log_{int(time.time())}.txt")
    success_log = []

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

        for index, row in df.iterrows():
            email = row["Email"]
            name = row["Name"] if has_name and pd.notna(row["Name"]) else "Sir/Mam"
            personalized_message = message_template.replace("(Name)", name)
            full_message = f"{personalized_message}\n\nBest regards,\nTeam Nutrition by Argha"

            msg = EmailMessage()
            msg["Subject"] = "Important Update from Nutrition by Argha"
            msg["From"] = EMAIL_ADDRESS
            msg["To"] = email
            msg.set_content(full_message)

            if attachment_path:
                with open(attachment_path, "rb") as f:
                    file_data = f.read()
                    file_name = os.path.basename(attachment_path)
                    maintype, subtype = ("application", "pdf") if file_name.endswith(".pdf") else ("image", "jpeg")
                    msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=file_name)

            try:
                smtp.send_message(msg)
                success_log.append(f"{email} - ✅ Sent")
            except Exception as e:
                success_log.append(f"{email} - ❌ Failed: {str(e)}")

            time.sleep(2)  # 2-second delay to reduce spam risk

    with open(log_file_path, "w") as log_file:
        log_file.write("\n".join(success_log))

    return jsonify({
        "status": "success",
        "log_url": f"/download/{os.path.basename(log_file_path)}"
    })

@app.route("/download/<filename>")
def download_log(filename):
    return send_from_directory(LOG_FOLDER, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
