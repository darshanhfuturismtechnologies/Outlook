
from Helper.helper import Helper

"""Tc_2:- Click on unread mail, click on most recent mail and read data and mark it as read"""

class UnreadMail(Helper):
    def __init__(self, app):
        super().__init__(app, title_re=".*Outlook.*")
        self.selected_mail = None

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
            self.selected_mail = recent_mail
            print(f"Most recent unread email clicked: {recent_mail}")
        else:
            raise Exception("No emails found in the inbox")

    def right_click_on_mail(self):
        try:
            self.selected_mail.right_click_input()
            self.logger.info("Right-clicked on the selected email")
        except:
            raise Exception("No email has been selected to right-click")


    def mark_as_read_or_unread(self):
        try:
            #Wait for the context menu or dropdown to become visible
            context_menu = self.outlook.child_window(title="Context Menu", control_type="Menu")
            context_menu.wait('visible', timeout=5)

            #First try Mark as read
            try:
                mark_as_read = context_menu.child_window(title="Mark as Read", control_type="MenuItem")
                if mark_as_read.exists(timeout=2):
                    mark_as_read.click_input()
                    self.logger.info("Email marked as read.")
                    return                                         #Exit after successful action
            except Exception as e:
                self.logger.debug(f"Mark as Read not found:{e}")

            #If Mark as Read not found then try Mark as unread
            try:
                mark_as_unread = context_menu.child_window(title="Mark as Unread", control_type="MenuItem")
                if mark_as_unread.exists(timeout=2):
                    mark_as_unread.click_input()
                    self.logger.info("Email marked as unread.")
                    return                                         #Exit after successful action
            except Exception as e:
                self.logger.debug(f"Mark as Unread not found:{e}")

        except Exception as e:
            self.logger.error(f"Failed to interact with context menu: {e}")








