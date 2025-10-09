import time
from Helper.helper import Helper

"""TC_05:-
1.Click on sent item
2.click on the mail that has subjcet as Demo Subject
3.Right click on that mail
4.Click on Delete option"""

class DeleteMail(Helper):
    def __init__(self,app):
        super().__init__(app, ".*Outlook.*")
        self.selected_mail = None

    def click_on_sent_item(self):
        sent_items=self.outlook.child_window(title="Sent Items", control_type="TreeItem",found_index=0).wait('visible',timeout=5)
        sent_items.click_input()

    def click_on_demo_sub_mail(self):
        #Parent controller of the mail
        group_box = self.outlook.child_window(title="Group By: Expanded: Date: Today", control_type="Group")
        #Children controllers which defines individual mails.
        mail_items = group_box.children(control_type="DataItem")

        #Target Subject to find
        target_sub = "Demo subject"

        #Find first email with target subject and click it
        for email in mail_items:
            if target_sub in email.window_text():
                email.click_input()
                time.sleep(1)                                                    #Wait for email visibility
                print(f"Clicked email with subject: {target_sub}")
                #Email is stored as selected_mail for further use
                self.selected_mail=email
                #Exit the loop after clicking
                break
        else:
            #If email is not clicked raise error
            raise Exception(f"No email found with subject: {target_sub}")

    def right_click_on_mail(self):
        try:
            self.selected_mail.right_click_input()
            self.logger.info("Right-clicked on the selected email")
        except:
            raise Exception("No email has been selected to right-click")

    def click_on_delete_mail(self):
        # Wait for the right click menu/context menu options to visible
        context_menu = self.outlook.child_window(title="Context Menu", control_type="Menu")
        context_menu.wait('visible', timeout=5)
        dlt_btn = context_menu.child_window(title="Delete", control_type="MenuItem")
        dlt_btn.click_input()
        self.logger.info("Clicked on Delete")


























