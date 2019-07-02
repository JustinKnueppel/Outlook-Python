from typing import List
from .email import Email

from itertools import compress

def search(emails: List[Email], sender_names: List[str]=[], sender_emails: List[str]=[], subjects: List[str]=[], bodies: List[str]=[], attachments: List[str]=[]) -> List[Email]:
    """Return a subset of emails containing any of the given criteria"""
    matches = [False]*len(emails)

    if sender_names:
        # Check if any of the given sender names appear
        sender_name_matches = [any(sender.lower() in email.sender_name().lower() for sender in sender_names) for email in emails]
        matches = [x or y for x, y in zip(matches, sender_name_matches)]

    if sender_emails:
        # Check if any of the given sender emails appear
        sender_email_matches = [any(sender.lower() in email.sender_email().lower() for sender in sender_emails) for email in emails]
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
        # Check if any of the given attachment names appear in any attachments
        attachment_matches = [any(any(attachment_target.lower() in attachment.name().lower() for attachment_target in attachments) for attachment in email.attachments()) for email in emails]
        matches = [x or y for x, y in zip(matches, attachment_matches)]

    return list(compress(emails, matches))
