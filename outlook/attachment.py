import win32com
import win32com.client

from typing import List

class Attachment:
    def __init__(self, win32_attachment: win32com.client.CDispatch):
        self._attachment = win32_attachment

    def __repr__(self):
        return self._attachment.FileName

    def __str__(self):
        return self._attachment.FileName

    def name(self) -> str:
        """Return the name of the attachment"""
        return self._attachment.FileName

    def extension(self) -> str:
        """Return the file extension of the attachment"""
        if '.' not in self.name():
            return ''

        return self.name().split('.')[-1]

    def save(self, filepath: str) -> bool:
        """Save file to the given filepath"""
        try:
            self._attachment.SaveASFile(filepath)
            return True
        except OSError:
            return False
