import win32com
import win32com.client
from datetime import datetime

from typing import List

class Email:
    def __init__(self, win32_email: win32com.client.CDispatch):
        self._email = win32_email

    def __repr__(self):
        return __email_to_string(self)

    def __str__(self):
        return __email_to_string(self)

    def attachments(self) -> List[win32com.client.CDispatch]:
        """Return a list of attachments"""
        num_attachments = self._email.Attachments.Count
        return [self._email.Attachments.Item(i) for i in range(num_attachments)]

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

def __email_to_string(email: Email) -> str:
    """Return string representation of an email object"""
    return f'{email.sender_name}: {email.subject}'
