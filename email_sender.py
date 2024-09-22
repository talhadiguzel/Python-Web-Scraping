import smtplib
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from datetime import datetime
import openpyxl

load_dotenv()

def send_email(subject, body, attachment_path,tomail):
    try:
        from_email = os.getenv("EMAIL_USER")
        password = os.getenv("PASSWORD")
        toemails = tomail

        if from_email is None:
            raise ValueError("EMAIL_USER environment variable is not set.")
        if password is None:
            raise ValueError("PASSWORD environment variable is not set.")
        if toemails is None:
            raise ValueError("TO_EMAIL environment variable is not set.")

        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(from_email, password)

        for toemail in toemails:
            msg = MIMEMultipart()
            msg['From'] = from_email
            msg['To'] = toemail
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))

            filename = os.path.basename(attachment_path)
            attachment = open(attachment_path, "rb")

            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename= {filename}')

            msg.attach(part)

            server.sendmail(from_email, toemail, msg.as_string())
            print(f"Email başarıyla gönderildi: {toemail}")

            wb = openpyxl.load_workbook(r"./mail_gonderim.xlsx")
            ws = wb.active

            line = ws.max_row + 1

            current_date = datetime.now().strftime("%Y-%m-%d")
            current_time = datetime.now().strftime("%H:%M:%S")
            mail_info = [from_email, toemail, current_date, current_time]
            ws.append(mail_info)
            wb.save(r"./mail_gonderim.xlsx")
            print(f"Email gönderim bilgileri kaydedildi: {toemail}")

        server.quit()

    except Exception as e:
        print(f"Email gönderim hatası: {e}")