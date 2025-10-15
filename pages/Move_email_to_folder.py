import time
from Helper.helper import Helper

class MoveToArchiveFolder(Helper):
    def __init__(self,app):
        super().__init__(app, ".*Outlook.*")         #Super constructor overrides properties of helper class constructor
        self.selected_mail = None

    def click_on_inbox(self):
        inbox_menu = self.outlook.child_window(title_re=".*Inbox.*", control_type="TreeItem",found_index=0).wait('visible',timeout=5)
        inbox_menu.click_input()
        # self.outlook.print_control_identifiers()

    def search_mail(self):
        try:
            search_box = self.outlook.child_window(auto_id="SearchBoxTextBoxAutomationId", control_type="Edit")
            search_box.type_keys("Demo Subject", with_spaces=True)
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
                self.selected_mail=mail_to_select                       #Assign selected mail for further use
                self.logger.info(f"Selected email with subject: {subject}")
        except Exception as e:
                self.logger.error(f"Error occurred: {e}")

    def right_click_on_mail(self):
        try:
            self.selected_mail.right_click_input()
            self.logger.info("Right-clicked on the selected email")
        except:
            raise Exception("No email has been selected to right-click")

    def click_on_move(self):
        # Wait for the right click menu/context menu options to visible
        context_menu = self.outlook.child_window(title="Context Menu", control_type="Menu")
        context_menu.wait('visible', timeout=5)
        dlt_btn = context_menu.child_window(title="Move", control_type="MenuItem")
        dlt_btn.click_input()
        self.logger.info("Clicked on Move")

    def click_on_archive(self):
        archive=self.outlook.child_window(title="Archive", control_type="ListItem")
        archive.click_input()
        self.logger.info("Clicked on Archive")



