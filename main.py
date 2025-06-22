import pandas as pd
import smtplib
import requests
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from jinja2 import Template
from dotenv import load_dotenv
import os

load_dotenv()  # Load email credentials

# Load email credentials
SENDER_EMAIL = os.getenv("EMAIL")
APP_PASSWORD = os.getenv("APP_PASSWORD")

# Load Excel file
data = pd.read_excel("application.xlsx")

# Load HTML template
with open("email_template.html", "r", encoding="utf-8") as f:
    template = Template(f.read())

# Load your resume PDF
resume_link = "https://drive.google.com/uc?export=download&id=17T6yeihhNswEHnqsLZHiVYFM_mXC62NK"
resume_bytes = requests.get(resume_link).content

# Placeholder photo
photo_url = "https://ui-avatars.com/api/?name=Tanmay+Binjola&background=random"

# SMTP setup
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(SENDER_EMAIL, APP_PASSWORD)

for index, row in data.iterrows():
    try:
        msg = MIMEMultipart()
        msg["From"] = SENDER_EMAIL
        msg["To"] = row["Email"]
        msg["Subject"] = f"Application for {row['Designation']} – Tanmay Binjola"

        html_body = template.render(
            photo_url=photo_url,
            company=row["Company"],
            designation=row["Designation"],
            company_logo=row["Logo"]
        )

        msg.attach(MIMEText(html_body, "html"))

        # Attach resume
        part = MIMEApplication(resume_bytes, Name="Tanmay_Binjola_CV.pdf")
        part['Content-Disposition'] = 'attachment; filename="Tanmay_Binjola_CV.pdf"'
        msg.attach(part)

        server.send_message(msg)
        print(f"[✓] Email sent to {row['Email']} at {row['Company']}")

    except Exception as e:
        print(f"[✗] Failed to send to {row['Email']}: {e}")

server.quit()
