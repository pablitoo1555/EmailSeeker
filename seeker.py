import os
from datetime import timedelta

from config import output_dir


class MsgParser:
    def __init__(self, contents, max_date, n_days):
        self.contents = contents
        self.path = os.getcwd()
        self.range = [max_date - timedelta(days=x) for x in range(n_days)]
        self.subject = None
        self.msg_count = 0

    def iso_msg(self, subject):
        """
        isolate message run parsing jobs
        :param subject: ( string ) expected subject name to search for (can be partial string)
        """
        try:
            print(f'searching for files between {min(self.range)} and {max(self.range)}...')
            for msg in self.contents.Items:
                self.subject = subject
                self._within_date_range(msg)

            return f'{self.msg_count} attachment(s) saved'
        except ValueError:
            raise ValueError('min() arg is an empty sequence (e.g. date range is none)')

    def _within_date_range(self, msg):
        """
        keep only messages(msg) where they meet the date range requirement
        :param msg: ( email object )
        """
        if msg.senton.date() in self.range:
            self._matches_subject_request(msg)
        else:
            pass

    def _matches_subject_request(self, msg):
        """
        check arg subject is in msg subject
        :param msg: ( email object )
        """
        if self.subject in msg.subject:
            self.save_attachment(msg)
        else:
            pass

    def save_msg(self, msg):
        """
        used to save email in .msg
        :param msg: ( email_obj )
        """
        msg.SaveAs(os.path.join(self.path, f'{"email" + "_" + str(msg.senton.date())}.msg'))

    def save_attachment(self, msg):
        """
        used to save attachment included in .msg
        :param msg: ( email obj )
        """
        for attachment in msg.Attachments:
            acct = msg.subject[0:6]
            date = msg.subject[-10:].replace("/", "-")
            file_name = f'{date}_COB_{acct}_Consolidated And Position Statement.pdf'
            attachment.SaveAsFile(os.path.join(self.path, output_dir, file_name))
            self.msg_count += 1
