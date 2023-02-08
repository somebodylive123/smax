import poplib
import smtplib
import mailparser
import base64
import os
from dotenv import load_dotenv
from bs4 import BeautifulSoup
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
load_dotenv()

# Connect to the mailbox
pop3_server = os.getenv('MAIL_RECIEVE_SERVER')
pop3_port = os.getenv('MAIL_RECIEVE_PORT')
pop3_ssl = os.getenv('MAIL_RECIEVE_SSL')
pop_conn = poplib.POP3_SSL(pop3_server, pop3_port)
pop_conn.user(os.getenv('WORK_MAIL'))
pop_conn.pass_(os.getenv('WORK_MAIL_PASSWORD'))



_, msg_list, _ = pop_conn.list()

print(msg_list)
# Sort the messages by index and retrieve the last message
if msg_list:
    msg_list = [int(x.split()[0]) for x in msg_list]
    msg_list.sort(reverse=True)
    msg_bytes = b"\n".join(pop_conn.retr(msg_list[0])[1])
    # original_email = message_from_bytes(msg_bytes)
    mail = mailparser.parse_from_bytes(msg_bytes)

    # encoded_subject = original_email['Subject'].encode('utf-8')
    # Create a new email to forward the original email
    forwarded_email = MIMEMultipart()
    forwarded_email['Subject'] = str(mail.subject)
    forwarded_email['From'] = (os.getenv('WORK_MAIL'))
    forwarded_email['To'] = (os.getenv('USER_MAIL'))
    forwarded_email['Date'] = str(mail.date)

    # Add the original email as an attachment
    body = mail.body
    soup = BeautifulSoup(body, 'html.parser')
    text = str(soup.get_text())
    text = text.strip()
    text_part = MIMEText(f"Отправитель: {str(mail.from_[0][0]), str(mail.from_[0][1])}\n Дата отправления: {str(mail.date)}\n Содержание: {str(text)}", 'plain', 'utf-8')
    forwarded_email.attach(text_part)

    # Add any attached files to the forwarded email
    for file in mail.attachments:
        file_data = file["payload"]
        base64.b64decode(file_data)
        attached_file = MIMEApplication(file_data, Name=str(file["filename"]))
        attached_file['Content-Disposition'] = 'attachment; filename="%s"' % str(file["filename"])
        forwarded_email.attach(attached_file)

    # Connect to the SMTP server
    smtp_server = os.getenv('MAIL_SEND_SERVER')
    smtp_port = os.getenv('MAIL_SEND_PORT')
    smtp_ssl = os.getenv('MAIL_SEND_SSL')
    smtp_conn = smtplib.SMTP(smtp_server, smtp_port)
    smtp_conn.starttls()
    smtp_conn.login((os.getenv('WORK_MAIL')), (os.getenv('WORK_MAIL_PASSWORD')))
    smtp_conn.sendmail((os.getenv('WORK_MAIL')), (os.getenv('USER_MAIL')), forwarded_email.as_string().encode("utf-8"))
    smtp_conn.quit()

    pop_conn.dele(msg_list[0])

# Close the connection to the mailbox
pop_conn.quit()
