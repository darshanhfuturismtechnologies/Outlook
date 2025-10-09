
from Helper.helper import Helper

"""Tc_2:- Click on unread mail, click on most recent mail and read data"""

class UnreadMail(Helper):
    def __init__(self, app):
        super().__init__(app, title_re=".*Outlook.*")

    def click_on_unread_mail(self):
        # Click on the 'Unread' text control to filter unread emails
        self.click(control_title="Unread", control_type="Text")
        # self.outlook.print_control_identifiers()

    def click_on_recent_unread_mail(self):
        # Find the group container with unread mails
        group = self.outlook.child_window(title="Group By: Expanded: Date: Older", control_type="Group")
        mail_items = group.children(control_type="DataItem")

        if mail_items:
            recent_mail = mail_items[0]
            recent_mail.click_input()
            print(f"Most recent unread email clicked: {recent_mail}")
        else:
            raise Exception("No emails found in the inbox")








