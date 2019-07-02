import win32com
import win32com.client
from datetime import datetime

from typing import List

from .attachment import Attachment

class Email:
    def __init__(self, win32_email: win32com.client.CDispatch):
        self._email = win32_email

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
        return self._email.RecievedTime

    def sender_email(self) -> str:
        """Returns the email address of the sender"""
        return str(self._email.SenderEmailAddress)

    def sender_name(self) -> str:
        """Returns the name of the sender"""
        return str(self._email.SenderName)

    def subject(self) -> str:
        """Returns the subject line of the email"""
        return str(self._email.Subject)
