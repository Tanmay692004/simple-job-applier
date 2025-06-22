# Simple Job applier

A Python project that sends personalized job application emails using an Excel file. Emails are sent with a branded HTML template and your resume as an attachment.

## 💡 Features
- Imports data from Excel: Email, Company, Designation, Logo
- Sends responsive HTML emails
- Attaches your resume (from Google Drive)
- Uses Gmail SMTP securely via .env

## 🔧 How to Use

1. Fill out `application.xlsx` with your contacts
2. Add your `.env` with Gmail and App Password
3. Run:

```bash
pip install -r requirements.txt
python main.py
