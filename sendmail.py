import logging
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

# setting up the logging object
logger = logging.getLogger()

# reading the config file
with open('mail_config.json', 'r', encoding='utf-8') as file:
    config = json.loads(file.read())

SENDER_USERNAME = config["sender_username"]
SENDER_PASSWORD = config["sender_password"]
SMTP_HOST = config["smtp_host"]
SMTP_PORT = config["smtp_port"]
RECEIVER_ADDRESS = config["receiver_address"]


def sending_mail(subject, message):

    """ A simple method that will send mails via SMTP """

    msg = MIMEMultipart()
    message = message
    msg['From'] = SENDER_USERNAME
    msg['To'] = ", ".join(RECEIVER_ADDRESS)
    msg['Subject'] = subject

    # add in the message body
    msg.attach(MIMEText(message, 'html'))

    #create server
    server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
    server.starttls()

    #Login
    server.login(msg['From'], SENDER_PASSWORD)

    # send the message via the server.
    server.sendmail(SENDER_USERNAME, RECEIVER_ADDRESS, msg.as_string())
    server.quit()
    logger.info("Mail sent to %s:" % (msg['To']))

