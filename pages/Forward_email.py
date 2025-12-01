import time
from Helper.helper import Helper

"""Tc_08
1.Click on Inbox.
2.Click on search bar the categorize menu will appear.
3.click on categorize menu and select Blue category.
4.Select recent email which is marked as blue category.
5.Click on forward.
6.Click on pop_out screen.
7.Enter details like To,Cc,Msg_body and click on Send button.
8.Clear search bar."""

class ForwardE(Helper):
    def __init__(self, app,test_data):
        super().__init__(app, ".*Outlook.*")
        self.FW_window = None
        self.test_data = test_data

    def click_on_inbox(self):
        self.click_menu_item(".*Inbox.*")
        self.logger.info("Clicked on Inbox")

    def select_most_recent_mail_of_demo_sub(self):
        try:
            self.outlook.child_window(auto_id="SearchBoxTextBoxAutomationId", control_type="Edit").click_input()
        except Exception as e:
            self.logger.error(f"Failed to click on Search bar:{e}")

    def click_on_categorized_menu_and_select_category(self):
        try:
            self.outlook.child_window(title="Categorized", control_type="MenuItem").click_input()
            self.logger.info("Categorized menu clicked")

            group_box = self.outlook.child_window(title=" ", control_type="Group", found_index=0).wait('visible',timeout=5)
            category_item = group_box.descendants(control_type="ListItem")

            for category in category_item:
                if category.window_text() == "Blue Category":
                    category.click_input()
                    time.sleep(1)
                    self.logger.info("Clicked on 'Blue Category'")
                    break
                else:
                    raise Exception("Blue Category not found.")
        except Exception as e:
            self.logger.error(f"Failed to select blue category item:{e}")
            raise

    def click_on_the_recent_mail(self):
        try:
            table_view = self.outlook.child_window(title="Table View", control_type="Table")
            mail_items = table_view.descendants(control_type="DataItem")
            self.logger.info(f"Total mails:{len(mail_items)}")

            if mail_items:
                recent_mail = mail_items[0]
                recent_mail.click_input()
                self.logger.info(f"Most recent email: {recent_mail}")
            else:
                raise Exception("No emails found in the inbox")
        except Exception as e:
            self.logger.error(f"Failed to click on recent mail:{e}")

    def click_on_forward_email(self):
        self.click(control_title="Forward", control_type="Button")

    def pop_out_screen(self):
        self.click(control_title="Pop Out", control_type="Button")
        self.logger.info("Clicked on Pop Out")

    def enter_details(self):
        try:
            self.FW_window = self.app.window(title_re=".*FW:.*")
            self.FW_window.wait('ready', timeout=5)
            self.logger.info("Forward window is visible")
            self.FW_window.set_focus()

            to_field = self.FW_window.child_window(auto_id="4117", control_type="Edit").wait('ready', timeout=5)
            to_field.type_keys(self.test_data["to"],pause=0.1, with_spaces=True)
            to_field.type_keys("{ENTER}")
            self.logger.info("to_field typed successfully")

            Cc = self.FW_window.child_window(auto_id="4126", control_type="Edit").wait('ready', timeout=5)
            Cc.type_keys(self.test_data["cc"],pause=0.1, with_spaces=True)
            Cc.type_keys("{ENTER}")
            self.logger.info("Cc typed successfully")

            msg_body = self.FW_window.child_window(title="Page 1 content", auto_id="Body", control_type="Edit")
            lines = self.test_data["body"].split("\n")
            for line in lines:
                # For loop iterate for each one and press enter twice,and add pause with spaces.
                msg_body.type_keys(line + "{ENTER 2}", pause=0.1, with_spaces=True)

        except Exception as e:
            self.logger.error(f"Failed to enter details:{e}")

    def click_on_send(self):
        try:
            send_btn = self.FW_window.child_window(title="Send", auto_id="4256", control_type="Button")
            send_btn.click_input()
            self.logger.info("Clicked on Send")
        except Exception as e:
            self.logger.error(f"Failed to click on send:{e}")
            raise
        self.wait_for_window_to_close(self.FW_window,timeout=30)

    def clear_search_bar(self):
        try:
            search_box = self.outlook.child_window(auto_id="SearchBoxTextBoxAutomationId", control_type="Edit").wait('ready',timeout=5)
            search_box.set_focus()
            search_box.type_keys("^a{BACKSPACE}")  # Clears the text
            time.sleep(1)
            self.logger.info("Search bar cleared.")
        except Exception as e:
            self.logger.error(f"Failed to clear search bar:{e}")








