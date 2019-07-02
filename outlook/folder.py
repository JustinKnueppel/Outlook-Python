import win32com
import win32com.client

from typing import List

from .email import Email
from .util import search

class Folder:
    def __init__(self, win32_folder: win32com.client.CDispatch):
        self._folder = win32_folder

    def __repr__(self):
        return self._folder.Name

    def __str__(self):
        return self._folder.Name

    def emails(self, recursive: bool=True) -> List[Email]:
        """Return collection of emails in the given folder"""
        return [Email(email) for email in self._folder.Items if email.MessageClass == 'IPM.Note']

    def folders(self) -> List['Folder']:
        """Return collection of subfolders in the given folder"""
        return [Folder(folder) for folder in self._folder.Folders]

    def search(self, recursive=False, **kwargs) -> List[Email]:
        """Return collection of emails matching the given search critera"""
        emails = self.emails()
        matches = search(emails, **kwargs)

        if recursive:
            for subfolder in self.folders():
                matches.extend(subfolder.search(recursive=True, **kwargs))
        
        return matches
