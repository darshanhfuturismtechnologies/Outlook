from Helper.helper import Helper

"""TC_01:-Click on inbox,click on recent mail and read the data"""

class ReadMail(Helper):
    def __init__(self,app):
        super().__init__(app, ".*Outlook.*")         #Super constructor overrides properties of helper class constructor

    def click_on_inbox(self):
        inbox_menu = self.outlook.child_window(title_re=".*Inbox.*", control_type="TreeItem",found_index=0).wait('visible',timeout=5)
        inbox_menu.click_input()
        # self.outlook.print_control_identifiers()

    def select_most_recent_email(self):
        #Find the parent container that holds the list of emails.
        group_box = self.outlook.child_window(title="Group By: Expanded: Date: Today", control_type="Group")
        #Find all the children DataItem controls which represent individual emails
        mail_items = group_box.children(control_type="DataItem")

        if mail_items:
            #Assuming the first mail is the most recent
            recent_mail = mail_items[0]
            # print(f"Most recent email: {recent_mail}")
            recent_mail.click_input()
            print(f"Most recent email: {recent_mail}")
        else:
            raise Exception("No emails found in the inbox")












