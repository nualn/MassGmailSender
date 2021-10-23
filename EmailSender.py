
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
import mailService

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

    # Sending mails with a loop.
    i = 1

    while sheet.cell_value(i, 0) != "":
        print("Sending " + sheet.cell_value(i, 1) + " to " + sheet.cell_value(i, 0))
        mailService.send_mail(
            config.SENDER, 
            sheet.cell_value(i, 0), 
            config.SUBJECT,
            config.MSG_BODY,
            sheet.cell_value(i, 1)
        )
        i += 1
    

    print("Sent " + str(i - 1) + " mails.")


