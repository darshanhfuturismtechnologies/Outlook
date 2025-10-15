from Helper.helper import Helper

"""TC_01:-Click on inbox,click on recent mail and read the data,make right click and mark it as read/unread"""

class ReadMail(Helper):
    def __init__(self,app):
        super().__init__(app, ".*Outlook.*")         #Super constructor overrides properties of helper class constructor
        self.selected_mail = None

    def click_on_inbox(self):
        inbox_menu = self.outlook.child_window(title_re=".*Inbox.*", control_type="TreeItem",found_index=0).wait('visible',timeout=5)
        inbox_menu.click_input()
        # self.outlook.print_control_identifiers()

    def select_most_recent_email(self):
        # Get the container that holds all groups of emails
        table_view = self.outlook.child_window(title="Table View", control_type="Table")
        # Find all DataItem controls under that table — each representing an email
        mail_items = table_view.descendants(control_type="DataItem")

        # #Find the parent container that holds the list of emails.
        # group_box = self.outlook.child_window(title="Group By: Expanded: Date: Today", control_type="Group")
        # #Find all the children DataItem controls which represent individual emails
        # mail_items = group_box.children(control_type="DataItem")
        self.logger.info(f"Found {len(mail_items)} mail items")

        if mail_items:
            #Assuming the first mail is the most recent
            recent_mail = mail_items[0]
            # print(f"Most recent email: {recent_mail}")
            recent_mail.click_input()
            self.selected_mail = recent_mail
            print(f"Most recent email: {recent_mail}")
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















