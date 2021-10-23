import config

import base64
import imaplib
import json
import smtplib
import urllib.parse
import urllib.request
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Added these 
from email import encoders
from email.mime.base import MIMEBase
import xlrd
import oauth

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
    



if __name__ == '__main__':
    if config.GOOGLE_REFRESH_TOKEN is None:
        print('No refresh token found, obtaining one')
        refresh_token, access_token, expires_in = oauth.get_authorization(config.GOOGLE_CLIENT_ID, config.GOOGLE_CLIENT_SECRET)
        print('Set the following as your GOOGLE_REFRESH_TOKEN:', refresh_token)
        exit()

    # Replacing this with my own mail 
    """
    send_mail('skipolin.ensilumi@gmail.com', 'nuuttinikkola1@gmail.com',
              'A mail from you from Python',
              '<b>A mail from you from Python</b><br><br>' +
              'So happy to hear from you!')
    """

    # Open the excel file and read it
    wb = xlrd.open_workbook(config.XLS_LOC)
    sheet = wb.sheet_by_index(0)

    ### Moved connection opening calls here so they're not called for each mail.
    
    ###
    # Sending mails with a loop.
    i = 1

    while sheet.cell_value(i, 0) != "":
        print("Sending " + sheet.cell_value(i, 1) + " to " + sheet.cell_value(i, 0))
        send_mail(
            config.SENDER, 
            sheet.cell_value(i, 0), 
            config.SUBJECT,
            config.MSG_BODY,
            sheet.cell_value(i, 1)
        )
        i += 1
    

    print("Sent " + str(i - 1) + " mails.")