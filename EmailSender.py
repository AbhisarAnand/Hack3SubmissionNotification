import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

from constants import *


class EmailSender:

    def __init__(self, email_address: str, time_slot: str, judging_group: str, team_member_names: str):
        self.receiver_email = email_address
        self.sender_email = SENDING_EMAIL_ADDRESS
        self.password = PASSWORD
        self.time_slot = time_slot
        self.judging_group = judging_group
        self.team_member_names = team_member_names
        self.zoom_link = ZOOM_LINK
        self.status = "fail"

    def send_email(self):
        msg = MIMEMultipart()
        sender_email = self.sender_email
        receiver_email = self.receiver_email
        password = self.password
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = "Hack3 Judging Time Slot for presentation"
        body = f"Hello {self.team_member_names},\n\nYour team has been assigned to {self.judging_group}, and your " \
               f"designated time slot is at {self.time_slot}.\n\n5 minutes before your designated time slot, " \
               f"please join the zoom link that is attached below.\n\n{self.zoom_link}\n\nIf the time selected " \
               f"doesn't work for you, DM an organizer in Slack ASAP.\n\nThank you,\nHack3 Team "
        msg.attach(MIMEText(body, "plain"))
        context = ssl.create_default_context()
        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                server.login(sender_email, password)
                server.sendmail(
                    sender_email, receiver_email.split(","), msg.as_string())
            self.status = "success"
        except Exception as e:
            print("RECEIVING EMAIL:" + self.receiver_email)
            print(type(e).__name__ + ': ' + str(e))
            self.status = "fail"
