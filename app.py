import win32com.client
import win32com
import os
import glob
from datetime import date

from seeker import MsgParser
from config import subject, trg_folder, latest_date, output_dir

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts


def _latest_run():
    """
    find the latest file modification date in output dir.
    Used to determine if files are missing and establish search range.
    :return: ( int ) days since last modification date and today
    """
    try:
        file_list = glob.glob(os.path.join(output_dir, '*'))
        last_file = max(file_list, key=os.path.getmtime)
        mod_time = date.fromtimestamp(os.path.getmtime(last_file))
        return abs(mod_time - date.today()).days - 1
    except ValueError:
        print(ValueError('max() arg is an empty sequence (e.g. output directory empty)'))
        return 1000


def run_app(n_days):
    """
    run main application
    :param n_days: ( int ) days since last assumed run
    """
    for account in accounts:
        inbox = outlook.Folders(account.DeliveryStore.DisplayName)
        target_folder = inbox.Folders['Inbox'].Folders[f'{trg_folder}']
        msg_parser_obj = MsgParser(target_folder, latest_date, n_days)
        print(msg_parser_obj.iso_msg(subject))


if __name__ == '__main__':
    n = _latest_run()
    run_app(n)
