import time
from Helper.helper import Helper

"""Tc_13:
1.Click on inbox.
2.Click on search bar & search most recent mail by text.
3.Right click on that mail and mark it as flagged.
4.Click on inbox & create one folder.
5.Click on that flagged mail and move it in newly created folder.
6.Delete the folder.
"""

class FlaggedMails(Helper):
    def __init__(self, app,test_data):
        super().__init__(app, ".*Outlook.*")
        self.test_data = test_data

    def click_on_inbox(self):
        self.click_menu_item(".*Inbox.*")

    def search_mail_of_demo_sub(self):
        try:
            search_box = self.outlook.child_window(auto_id="SearchBoxTextBoxAutomationId", control_type="Edit")
            search_box.type_keys(self.test_data["search_box_text"], with_spaces=True)
            time.sleep(1)

            table_view = self.outlook.child_window(title="Table View", control_type="Table")
            mail_items = table_view.descendants(control_type="DataItem")
            self.logger.info(f"Found {len(mail_items)} mail items")

            if len(mail_items) > 0:
                mail_to_select = mail_items[0]
                subject = mail_to_select.window_text()    #Capture subject before double-click
                mail_to_select.click_input()
                self.selected_mail = mail_to_select       #Assign selected mail for further use
                self.logger.info(f"Selected email with subject:{subject}")
        except Exception as e:
            self.logger.error(f"Error occurred:{e}")

    def right_click_on_mail(self):
        self.right_click_on_selected_option()

    def click_on_follow_up(self):
        try:
            context_menu=self.get_context_menu()
            dlt_btn = context_menu.child_window(title="Follow Up",control_type="MenuItem")
            dlt_btn.click_input()
            self.logger.info("Clicked on Follow Up")
        except Exception as e:
            self.logger.error(f"Failed to click on 'Move' menu item: {e}")
            raise

    def click_on_no_date_flag(self):
        try:
            follow_up_menubar = self.outlook.child_window(title="Follow Up",control_type="Menu")
            submenus = follow_up_menubar.descendants(control_type="MenuItem")
            self.logger.info(f"Found {len(submenus)} sub menu items")

            for submenu in submenus:
                if submenu.window_text().strip()=="No Date".strip():
                    submenu.click_input()
                    self.logger.info(f"Clicked on No Date flag:{submenu.window_text()}")
                    break
        except Exception as e:
            self.logger.error(f"Failed to click on 'No Date' flag:{e}")

    def click_on_inbox_and_make_right_click (self):
        try:
            draft = self.outlook.child_window(title_re=".*Inbox.*", control_type="TreeItem")  # adjust type if needed
            draft.wait('ready',timeout=5)
            self.logger.info("Clicked on inbox menu")
            draft.right_click_input()
            self.logger.info("Right-clicked on inbox menu")
        except Exception as e:
            self.logger.error(f"Failed to click on inbox menu item: {e}")

    def click_on_create_new_folder(self):
        try:
            context_menu=self.get_context_menu()
            context_menu.child_window(title="New Folder...", control_type="MenuItem").click_input()
            self.logger.info("Clicked on Create New Folder menu item")
        except Exception as e:
            self.logger.error(f"Failed to click on 'Create New Folder' menu item:{e}")

    def enter_folder_name(self):
        try:
            sent_items = self.outlook.child_window(title_re=".*Create New Folder.*")
            sent_items.wait('visible', timeout=10)
            self.logger.info("Create New Folder window is visible")

            edit_name = sent_items.child_window(auto_id="4097",control_type="Edit")
            edit_name.wait('ready',timeout=5)
            edit_name.type_keys(self.test_data["folder_name"]+"{ENTER}",with_spaces=True,pause=0.1)  #press Enter to create
            self.logger.info(f"{self.test_data["folder_name"]} is created")
        except Exception as e:
             self.logger.error(f"Failed to create folder:{e}")

    def click_on_demo_folder_via_move(self):
        try:
            #Select the flagged mail
            self.selected_mail.right_click_input()
            time.sleep(1)
            self.logger.info("Right clicked on selected mail")

            #Open the Move menu
            move_menu = self.outlook.child_window(title="Move", control_type="MenuItem")
            move_menu.click_input()
            self.logger.info("Opened Move menu")
            time.sleep(1.5)

            #Access the nested Move menu
            move_submenu = self.outlook.child_window(title="Move", control_type="Menu", found_index=0)
            move_submenu.wait('visible',timeout=5)
            self.logger.info("Located Move submenu")

            #Inside that menu,find the GroupBox
            grp_box = move_submenu.child_window(title=" ", control_type="Group", found_index=0)
            grp_box.wait('visible',timeout=5)
            self.logger.info("Found GroupBox inside Move submenu")

            #Click the Demo folder ListItem
            demo_folder_item = grp_box.child_window(title_re=".*Demo folder.*", control_type="ListItem")
            demo_folder_item.wait('ready',timeout=5)
            demo_folder_item.click_input()
            # time.sleep(1)
            self.logger.info("Successfully moved email to Demo folder")

        except Exception as e:
            self.logger.error(f"Failed to move email to Demo folder via Move menu: {e}")
            raise

    def delete_demo_folder(self):
        try:
            #Select demo folder and right click on it to delete
            demo_folder = self.outlook.child_window(title="Demo folder", control_type="TreeItem")
            demo_folder.right_click_input()

            context_menu = self.get_context_menu()
            dlt_folder = context_menu.child_window(title="Delete Folder", control_type="MenuItem")
            dlt_folder.wait('ready',timeout=5)
            dlt_folder.click_input()
            self.logger.info("Deleted folder successfully")
        except Exception as e:
            self.logger.error(f"Failed to delete folder:{e}")

    def handle_pop_up(self):
        try:
            #Wait for the Delete Folder popup to appear
            dlt_pop_up = self.outlook.child_window(title="Microsoft Outlook",control_type="Window")
            dlt_pop_up.wait('visible', timeout=5)
            self.logger.info("Delete pop-up appeared.")

            #Check if the Yes button exists and click it.
            yes_button = dlt_pop_up.child_window(title="Yes", auto_id="6",control_type="Button")
            if yes_button.exists(timeout=2):
                yes_button.click_input()
                time.sleep(2)
                self.logger.info("Clicked Yes to delete the folder.")
            else:
                #Click No if Yes is not found
                no_button = dlt_pop_up.child_window(title="No", auto_id="2",control_type="Button")
                if no_button.exists(timeout=2):
                    no_button.click_input()
                    time.sleep(2)
                    self.logger.info("Clicked No in delete popup.")
                else:
                    self.logger.warning("Both Yes or No buttons were found in the popup.")

        except Exception as e:
            self.logger.error(f"Failed to handle delete popup:{e}")



























































