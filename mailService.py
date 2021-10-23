
import config
import oauth

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def send_mail(fromaddr, toaddr, subject, body, filename):
    access_token, expires_in = oauth.refresh_authorization(config.GOOGLE_CLIENT_ID, config.GOOGLE_CLIENT_SECRET, config.GOOGLE_REFRESH_TOKEN)
    auth_string = oauth.generate_oauth2_string(config.SENDER, access_token, as_base64=True)
    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = fromaddr
    message["To"] = toaddr
    message["Subject"] = subject

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    if filename:
        filepath = config.FILE_FOLDER + filename  # In same directory as script

        # Open PDF file in binary mode
        with open(filepath, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        # Encode file in ASCII characters to send by email    
        encoders.encode_base64(part)

        # Add header as key/value pair to attachment part
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filepath}",
        )

        # Add attachment to message and convert message to string
        message.attach(part)
        
    text = message.as_string()
    
    server = smtplib.SMTP('smtp.gmail.com:587')
    server.ehlo(config.GOOGLE_CLIENT_ID)
    server.starttls()
    server.docmd('AUTH', 'XOAUTH2 ' + auth_string)
    
    server.sendmail(fromaddr, toaddr, text)
    # Close server
    server.quit()
    
