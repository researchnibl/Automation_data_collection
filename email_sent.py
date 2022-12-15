import smtplib
from email.message import EmailMessage

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.audio import MIMEAudio
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders

from datetime import date

# set your email and password
# please use App Password
def email_sent():
    email_address = "tiwarilaxuu@gmail.com"
    email_password = "nxalheglkybmzdls"

    # create email
    # msg = EmailMessage()
    msg = MIMEMultipart()
    msg['Subject'] = "Daily Report Sent"
    msg['From'] = 'tiwarilaxuu@gmail.com'
    msg['To'] = "research.niblace@gmail.com"
    body = f"Daily Broker Report {date.today()}. This is automatic sent from Python  "

    msg.attach(MIMEText(body, 'plain'))

    filename = f"Top 7 broker List weekly {date.today()}.xlsx"
    attachment = open(f"Daily_Report/Top 7 broker List weekly {date.today()}.xlsx", "rb")

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    msg.attach(part)


    # send email
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(email_address, email_password)
        smtp.send_message(msg)
    print('Email send successfully')