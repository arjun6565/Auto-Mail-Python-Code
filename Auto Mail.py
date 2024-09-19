import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Sender's email credentials
from_email = "abc@gmail.com"
password = "abcdefghijk"  # Your Gmail app password

# Recipient's email ID
to_email = "abc@gmail.com"

# Email details
subject = "Auto MAil Daily Report with Excel Attachment"
body = "Please find attached the daily report.\n\nNote: this is an automatic mail, please do not reply."

# Attachment file details
attachment_path = "D:/Automation Testing/Automail.xlsx"

# Create a multipart message
msg = MIMEMultipart()
msg['From'] = from_email
msg['To'] = to_email
msg['Subject'] = subject

# Add message body to email
msg.attach(MIMEText(body, 'plain'))

# Add attachment to email
with open(attachment_path, "rb") as attachment:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {'Automail.xlsx'}",
    )
    msg.attach(part)

# Send the email
try:
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_email, password)
    text = msg.as_string()
    server.sendmail(from_email, to_email, text)
    server.quit()
    print("Sent message successfully with Excel attachment...")
except Exception as e:
    print("Error:", str(e))
