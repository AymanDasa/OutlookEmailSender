import win32com.client
ol = win32com.client.Dispatch('Outlook.Application')
olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)
attach = 'C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
newmail.Attachments.Add(attach)
newmail.Subject = 'Testing Mail'
newmail.To = 'to_mail@gmail.com'
newmail.CC = 'cc_mail@gmail.com' 
newmail.BCC = 'bcc_mail@gmail.com'
newmail.Body= 'Hello, this is a test email to showcase how to send emails from Python and Outlook.'
newmail.Display()
newmail.Send()
 
