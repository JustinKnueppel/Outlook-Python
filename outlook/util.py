from typing import List
from .email import Email

from itertools import compress

def search(emails: List[Email], senders: List[str]=[], subjects: List[str]=[], bodies: List[str]=[], attachments: List[str]=[]) -> List[Email]:
    """Return a subset of emails containing any of the given criteria"""
    matches = [False]*len(emails)

    if senders:
        # Check if any of the given senders appear
        sender_matches = [any(sender in email.sender() for sender in senders) for email in emails]
        matches = [x or y for x, y in zip(matches, sender_matches)]

    if subjects:
        # Check if any of the given subjects appear
        subject_matches = [any(subject in email.subject() for subject in subjects) for email in emails]
        matches = [x or y for x, y in zip(matches, subject_matches)]

    if bodies:
        # Check if any of the given bodies appear
        body_matches = [any(body in email.body() for body in bodies) for email in emails]
        matches = [x or y for x, y in zip(matches, body_matches)]

    if attachments:
        # Check if any of the given attachment names appear in any attachments
        attachment_matches = [any(any(attachment_target in attachment.name() for attachment_target in attachments) for attachment in email.attachments()) for email in emails]
        matches = [x or y for x, y in zip(matches, attachment_matches)]

    return list(compress(emails, matches))
