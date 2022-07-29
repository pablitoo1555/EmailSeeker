import os

from helpers import create_dir
from config import output_dir

class MsgParser:
    def __init__(self, subject):
        self.subject = subject
        self.path = os.getcwd()

    def iso_msg(self, msg):
        create_dir(output_dir)
        if msg.subject == self.subject:
            self.save_attachment(msg)
            print('Message Found')

    def save_msg(self, msg):
        msg.SaveAs(os.path.join(self.path, f'{"email" + "_" + str(msg.senton.date())}.msg'))

    def save_attachment(self, msg):
        """
        used to save attachment included in .msg
        :param msg:
        :return:
        """
        for attachment in msg.Attachments:
            attachment.SaveAsFile(os.path.join(self.path, output_dir, attachment.FileName))

