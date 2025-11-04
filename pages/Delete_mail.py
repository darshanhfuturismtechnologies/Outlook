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
        self.click_menu_item(".*Sent Items.*")
        self.logger.info(f"Successfully clicked on sent item.")


    def click_on_demo_sub_mail(self):
        # Get the container that holds all groups of emails
        table_view = self.outlook.child_window(title="Table View", control_type="Table")
        # Find all DataItem controls under that table — each representing an email
        mail_items = table_view.descendants(control_type="DataItem")

        #Target Subject to find
        target_sub = "Demo subject"

        #Find first email with target subject and click it
        for email in mail_items:
            if target_sub in email.window_text():
                email.click_input()
                time.sleep(1)                                                    #Wait for email visibility
                self.logger.info(f"Clicked email with subject: {target_sub}")
                #Email is stored as selected_mail for further use
                self.selected_mail=email
                #Exit the loop after clicking
                break
        else:
            #If email is not clicked raise error
            raise Exception(f"No email found with subject: {target_sub}")

    def right_click_on_mail(self):
        self.right_click_on_selected_option()

    def click_on_delete_mail(self):
        try:
            # Wait for the right click menu/context menu options to visible
            context_menu = self.outlook.child_window(title="Context Menu", control_type="Menu")
            context_menu.wait('visible', timeout=5)

            dlt_btn = context_menu.child_window(title="Delete", control_type="MenuItem")
            dlt_btn.click_input()
            self.logger.info("Clicked on Delete")

        except Exception as e:
            self.logger.error(f"Failed to click on Delete mail: {e}")
            raise































