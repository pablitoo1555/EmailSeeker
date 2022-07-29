import win32com.client
import win32com

from seeker import MsgParser
from config import subject, trg_folder

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts

MsgParser_obj = MsgParser(subject)


def run_app():
    for account in accounts:
        inbox = outlook.Folders(account.DeliveryStore.DisplayName)
        target_folder = inbox.Folders['Inbox'].Folders[f'{trg_folder}']
        for msg in target_folder.Items:
            MsgParser_obj.iso_msg(msg)


if __name__ == '__main__':
    run_app()
