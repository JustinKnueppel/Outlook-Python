import win32com
import win32com.client
from itertools import compress

from typing import List

from .email import Email

class Folder:
    def __init__(self, win32_folder: win32com.client.CDispatch):
        self._folder = win32_folder

    def __repr__(self):
        return self._folder.Name

    def __str__(self):
        return self._folder.Name

    def emails(self, recursive: bool=False) -> List[Email]:
        """Return collection of emails in the given folder"""
        emails = [Email(email) for email in self._folder.Items if email.MessageClass == 'IPM.Note']
        if recursive:
            for folder in self.folders():
                emails.extend(folder.emails(recursive=True))
        return emails

    def folders(self) -> List['Folder']:
        """Return collection of subfolders in the given folder"""
        return [Folder(folder) for folder in self._folder.Folders]
