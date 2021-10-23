# MassGmailSender

## Description

A simple script to send mass emails with custom attachments.

## Getting started
### Installation
Clone the repository:
```
git clone https://github.com/nikkoln/MassGmailSender.git
```
Create a config.py file in the root folder.
Add to it the following variables:
```
# Sender's email address
SENDER = ''
# Subject of the email
SUBJECT = ''
# Body of the email
MSG_BODY = ''
# Folder with attachments
FILE_FOLDER = ''
# Workbook that connects the mail addresses to files 
XLS_LOC = ''

# Parameters from Google's API development
GOOGLE_CLIENT_ID = ''
GOOGLE_CLIENT_SECRET = ''
# Token from when you first run the script
GOOGLE_REFRESH_TOKEN = None
```
You can get the client id and client secret from [here.](https://console.cloud.google.com/apis/)

You also need a xls worksheet with the following structure:
|email            |filename   |
|-----------------|-----------|
|example@email.com|example.pdf|
|and so           |on         |

Run main.py and follow the instructions printed in the console.
The second time you run main.py it will send emails according to the variables you have defined, to the emails in the worksheet, with the files from the worksheet (provided they exist in the folder you have determined) as attatchments. 


