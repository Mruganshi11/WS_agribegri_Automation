
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

def send_email_with_attachment(sender_email, sender_password, receiver_email, subject, body, attachment_path):
    try:
        if not receiver_email:
            print(" No receiver email provided. Skipping email.")
            return

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))

        # Attachment
        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {os.path.basename(attachment_path)}",
            )
            msg.attach(part)
        else:
            print(f" Attachment not found: {attachment_path}")
            return

        # SMTP Server (Assuming Gmail or standard SMTP, adjust as needed)
        # Use simple SMTP for now, or assume the user has a local relay or specific config
        # Since user said "Using placeholder", I will assume standard Gmail port 587
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        text = msg.as_string()
        server.sendmail(sender_email, receiver_email, text)
        server.quit()
        
        print(f" Email sent successfully to {receiver_email}")
    except Exception as e:
        print(f" Failed to send email: {e}")
        
def send_email_multiple_attachments(sender, password, receiver, subject, body, file_paths):
    from email.message import EmailMessage
    import smtplib
    import os

    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = receiver
    msg["Subject"] = subject
    msg.set_content(body)

    for path in file_paths:
        with open(path, "rb") as f:
            file_data = f.read()
            file_name = os.path.basename(path)

        msg.add_attachment(
            file_data,
            maintype="application",
            subtype="pdf",
            filename=file_name
        )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(sender, password)
        smtp.send_message(msg)
