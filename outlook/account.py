import win32com
import win32com.client

from typing import List

class Account:
    def __init__(self, account_name: str):
        assert any([account_name in account_display_name for account_display_name in Account.list_accounts()]), "Given account not found"

        self.outlook = win32com.client.Dispatch('Outlook.Application').GetNameSpace('MAPI')
        self.account = Account._get_account_by_name(account_name)

    @staticmethod
    def _get_account_by_name(account_name: str) -> win32com.client.CDispatch:
        """Retrieve an account based on the associated email"""
        account = next(filter(lambda account: account_name in account.DisplayName, win32com.client.Dispatch("Outlook.Application").Session.Accounts))
        return account

    def folders(self) -> List[str]:
        """Return collection of folder names"""
        return [str(folder) for folder in self.outlook.Folders(self.account.DeliveryStore.DisplayName).Folders]

    def get_folder(self, folder_name: str) -> win32com.client.CDispatch:
        """Return a folder based on its name"""
        folder = next(filter(lambda folder: folder_name.lower() in str(folder).lower(), self.outlook.Folders(self.account.DeliveryStore.DisplayName).Folders))
        return folder

    @staticmethod
    def list_accounts() -> List[str]:
        """List all accounts in the current session"""
        num_accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts.Count
        accounts = [account.DisplayName for account in [win32com.client.Dispatch("Outlook.Application").Session.Accounts[i] for i in range(num_accounts)]]
        return accounts
