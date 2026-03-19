import win32com.client as win32
from pathlib import Path

def send_email_with_outlook(subject, body, to_list, cc_list=None, bcc_list=None, attachment_path=None):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.Subject = subject
    mail.Body = body
    mail.To = ";".join(to_list)

    if cc_list:
        mail.CC = ";".join(cc_list)

    if bcc_list:
        mail.BCC = ";".join(bcc_list)

    if attachment_path and Path(attachment_path).exists():
        mail.Attachments.Add(str(attachment_path))

    mail.Send()   # Use mail.Display() if you want to see before sending

    print("Email sent successfully via Outlook!")

# Example usage
send_email_with_outlook(
    subject="Test Email-2",
    body="Hi Team,\n\nPlease ignore this mail as well.\n\nRegards,\nSiddhant",
    to_list=["indiait@aboutsib.com"],
    cc_list=["indiait@aboutsib.com"],
   
)