import win32com
import win32com.client

from typing import List

from .email import Email

class Folder:
    def __init__(self, win32_folder: win32com.client.CDispatch):
        self._folder = win32_folder

    def __repr__(self):
        return self._folder.Name

    def __str__(self):
        return self._folder.Name

    def emails(self, recursive: bool=True) -> List[Email]:
        """Return collection of emails in the given folder"""
        #TODO Limit this to only MailItems
        return [Email(email) for email in self._folder.Items]

    def folders(self) -> List[win32com.client.CDispatch]:
        """Return collection of subfolders in the given folder"""
        return list(self._folder.Folders)
