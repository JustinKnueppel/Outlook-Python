from .account import Account
from .attachment import Attachment
from .email import Email
from .folder import Folder

import win32com
import win32ui
import os
from itertools import compress
from datetime import datetime

from typing import List

def is_running() -> bool:
    """Return true if Outlook is running"""
    try:
        win32ui.FindWindow(None, 'Microsoft Outlook')
        return True
    except win32ui.error:
        return False

def open():
    """Open Outlook"""
    os.startfile('outlook')

def search(emails: List[Email], sender_names: List[str]=[], sender_emails: List[str]=[], subjects: List[str]=[], bodies: List[str]=[], attachments: List[str]=[]) -> List[Email]:
    """Return subset of emails containing any of the given criteria"""
    matches = [False]*len(emails)

    if sender_names:
        # Check if any of the given sender names appear
        sender_name_matches = [any(sender.lower() in email.sender_name().lower() for sender in sender_names) for email in emails]
        matches = [x or y for x, y in zip(matches, sender_name_matches)]
    
    if sender_emails:
        # Check if any of the given sender emails appear
        sender_email_matches = [any(sender_email.lower() in email.sender_email().lower() for sender_email in sender_emails) for email in emails]
        matches = [x or y for x, y in zip(matches, sender_email_matches)]

    if subjects:
        # Check if any of the given subjects appear
        subject_matches = [any(subject.lower() in email.subject().lower() for subject in subjects) for email in emails]
        matches = [x or y for x, y in zip(matches, subject_matches)]

    if bodies:
        # Check if any of the given bodies appear
        body_matches = [any(body.lower() in email.body().lower() for body in bodies) for email in emails]
        matches = [x or y for x, y in zip(matches, body_matches)]

    if attachments:
        # Check if any of the given attachment names appear
        attachment_name_matches = [any(any(attachment_target.lower() in attachment.name().lower() for attachment_target in attachments) for attachment in email.attachments()) for email in emails]
        matches = [x or y for x, y in zip(matches, attachment_name_matches)]

    return list(compress(emails, matches))

def limit_between(emails: List[Email], start_time: datetime, end_time: datetime) -> List[Email]:
    """Return given emails that fall between the given start and end time"""
    return list(filter(lambda email: start_time <= email.received_time() <= end_time, emails))
