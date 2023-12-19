# OutlookEmailSender

## Overview

This Python script utilizes the `win32com.client` library to interact with Microsoft Outlook, allowing you to send emails programmatically. The script creates a new email, attaches a file, sets various email properties such as subject, recipients, and body, and then sends the email.


## Prerequisites

- Python installed on your system.
- `win32com.client` library. You can install it using:

```
pip install pywin32
```  
## Usage
1. Clone or download the script to your local machine.
```
git clone https://github.com/your-username/outlook-email-sender.git
```
2. Update the `attach` variable with the correct file path of the attachment you want to include in your email.
```
attach = 'C:\\path\\to\\your\\attachment.xlsx'
```
3. Modify the email properties (subject, recipients, CC, BCC, body) based on your requirements.
```
newmail.Subject = 'Your Subject'
newmail.To = 'to_mail@outlook.com'
newmail.CC = 'cc_mail@outlook.com'
newmail.BCC = 'bcc_mail@outlook.com'
newmail.Body = 'Your email body here.'
```

4. Save the changes.
5. Run the script.
```
python outlook_email_sender.py
```
This will open Outlook with a new email window populated with the specified details. Review the email and send it manually.

## Important Note
- Ensure that Outlook is installed on your machine and configured with an email account.
- Allow access for Python to interact with Outlook.

## Contributing
Feel free to contribute to this project by submitting issues or pull requests.
