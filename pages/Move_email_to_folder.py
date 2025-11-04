import time
from Helper.helper import Helper

"""Tc_07
1.Click on inbox.
2.Click on search bar and search the email of Demo subject and click on that
3.Make right click on selected mail
4.Click on move
5.Click on archive"""

class MoveToArchiveFolder(Helper):
    def __init__(self,app):
        super().__init__(app, ".*Outlook.*")         #Super constructor overrides properties of helper class constructor
        self.selected_mail = None

    def click_on_inbox(self):
        self.click_menu_item(".*Inbox.*")
        self.logger.info(f"Successfully clicked on inbox.")

    def search_mail(self):
        try:
            search_box = self.outlook.child_window(auto_id="SearchBoxTextBoxAutomationId",control_type="Edit")
            search_box.type_keys("Demo Subject",with_spaces=True)
            time.sleep(2)                                               #To load elements

            #Get the container that holds all groups of emails
            table_view = self.outlook.child_window(title="Table View", control_type="Table")
            # Find all DataItem controls under that table — each representing an email
            mail_items = table_view.descendants(control_type="DataItem")
            self.logger.info(f"Found {len(mail_items)} mail items")

            if len(mail_items) > 0:
                mail_to_select = mail_items[1]                         #Index 1 for the second mail item
                subject = mail_to_select.window_text()                 #Capture subject before double-click
                mail_to_select.click_input()
                self.selected_mail=mail_to_select                     #Assign selected mail for further use
                self.logger.info(f"Selected email with subject: {subject}")
        except Exception as e:
                self.logger.error(f"Error occurred: {e}")

    def right_click_on_mail(self):
        self.right_click_on_selected_option()

    def click_on_move(self):
        try:
            # Wait for the right click menu/context menu options to visible
            context_menu = self.outlook.child_window(title="Context Menu", control_type="Menu")
            context_menu.wait('visible', timeout=5)
            move_btn = context_menu.child_window(title="Move", control_type="MenuItem")
            move_btn.click_input()
            self.logger.info("Clicked on Move")
        except Exception as e:
            self.logger.error(f"Failed to click on 'Move' menu item: {e}")
            raise

    def click_on_archive(self):
        try:
            archive = self.outlook.child_window(title="Archive", control_type="ListItem")
            archive.click_input()
            self.logger.info("Clicked on Archive")
        except Exception as e:
            self.logger.error(f"Failed to click on 'Archive' menu item: {e}")
            raise



