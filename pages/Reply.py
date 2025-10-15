import time
from Helper.helper import Helper
"""
TC_04
1.Click on inbox,
2.Select the mail which has subject-Demo subject,
3.Click on reply,
4.Edit msg in reply body,
5.Click on send"""

class ReplyMail(Helper):
    def __init__(self, app):
        super().__init__(app, ".*Outlook.*")

    def navigate_towards_inbox(self):
        inbox= self.outlook.child_window(title_re=".*Inbox.*", control_type="TreeItem", found_index=0).wait('visible',timeout=5)
        inbox.click_input()

    def click_on_first_mail_of_demo_sub(self):
        # Get the container that holds all groups of emails
        table_view = self.outlook.child_window(title="Table View", control_type="Table")
        # Find all DataItem controls under that table — each representing an email
        mail_items = table_view.descendants(control_type="DataItem")

        # #Parent controller of the mail
        # group_box = self.outlook.child_window(title="Group By: Expanded: Date: Today", control_type="Group")
        # #Children controllers which defines individual mails.
        # mail_items = group_box.children(control_type="DataItem")

        #Target Subject to find
        target_sub = "Demo subject"

        #Find first email with target subject and click it
        for email in mail_items:
            if target_sub in email.window_text():
                email.click_input()
                time.sleep(1)                                              #Wait for email visibility
                print(f"Clicked email with subject: {target_sub}")
                #Exit the loop after clicking
                break
        else:
        #If email is not clicked raise error
            raise Exception(f"No email found with subject: {target_sub}")

    def click_on_reply(self):
        reply_btn=self.outlook.child_window(title="Reply", control_type="Button",found_index=0).wait('visible',timeout=5)
        reply_btn.click_input()

    def edit_mail_for_reply(self):
        msg_body = self.outlook.child_window(title="Page 1 content", auto_id="Body", control_type="Edit")
        #List of content for mail
        lines = [
            "Hello, Swapnali K.",
            "This is Reply Testcase..!",
            "Best Regards, Swapnali K",
        ]
        for line in lines:
            #For loop iterate for each one and press enter twice,and add pause with spaces.
            msg_body.type_keys(line + "{ENTER 2}", pause=0.1, with_spaces=True)

    def click_on_send(self):
        self.click_child_window(control_title="Send", control_type="Button")
        time.sleep(20)




















