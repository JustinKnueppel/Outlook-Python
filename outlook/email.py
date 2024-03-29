import win32com
import win32com.client
from datetime import datetime
import re

from typing import List

from .attachment import Attachment

class Email:
    def __init__(self, win32_email: win32com.client.CDispatch):
        self._email = win32_email

    @classmethod
    def new_email(cls, outlook: win32com.client.CDispatch, to: str, subject: str, body: str) -> 'Email':
        """Create an email object from scratch"""
        email = outlook.CreateItem(0)
        email.To = to
        email.Subject = subject
        email.Body = body
        return cls(email)

    def __repr__(self):
        return f'{self.sender_name()}: {self.subject()}'

    def __str__(self):
        return f'{self.sender_name()}: {self.subject()}'

    def attachments(self) -> List[win32com.client.CDispatch]:
        """Return a list of attachments"""
        num_attachments = self._email.Attachments.Count
        return [Attachment(self._email.Attachments.Item(i)) for i in range(1, num_attachments + 1)]

    def body(self) -> str:
        """Returns the body of the email"""
        return str(self._email.Body)

    def received_time(self) -> datetime:
        """Returns the time the email was received"""
        pywintime = str(self._email.ReceivedTime)
        time_regex = re.compile(r'(\d{4})-' # Year
                                r'(\d{2})-' # Month
                                r'(\d{2}) ' # Day
                                r'(\d{2}):' # Hours
                                r'(\d{2}):' # Minutes
                                r'(\d{2}).*'# Seconds
        )
        match = time_regex.match(pywintime)
        return datetime(*map(int, match.groups()))

    def send(self) -> bool:
        """Send the given email"""
        #TODO: Add asserts to have a receiving address and content
        self._email.Send()

    def sender_email(self) -> str:
        """Returns the email address of the sender"""
        return str(self._email.SenderEmailAddress)

    def sender_name(self) -> str:
        """Returns the name of the sender"""
        return str(self._email.SenderName)

    def subject(self) -> str:
        """Returns the subject line of the email"""
        return str(self._email.Subject)
