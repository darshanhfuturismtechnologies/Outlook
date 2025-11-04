import os
import time
from itertools import count
from Helper.helper import Helper
"""
Tc_14
1.Search mail by content in search box and open the mail.
2.Click on attachment option.
3.Click on Save as.
4.Select the folder to save the attachment.
5.Save file with unique name.(If saving same file then it not work)
6.Edit file name.
7.Click on save.
"""

class Download_attachments(Helper):
    def __init__(self,app):
        super().__init__(app,".*Outlook.*")
        self.filename = None
        self.save_attachment_dialog = None


    def search_mail_of_demo_sub(self):
        try:
            search_box = self.outlook.child_window(auto_id="SearchBoxTextBoxAutomationId", control_type="Edit")
            search_box.type_keys("This is the body of  the page", with_spaces=True)
            time.sleep(1)

            table_view = self.outlook.child_window(title="Table View", control_type="Table")
            mail_items = table_view.descendants(control_type="DataItem")
            self.logger.info(f"Found {len(mail_items)} mail items")

            if len(mail_items) > 0:
                mail_to_select = mail_items[0]
                mail_to_select.click_input()
                self.logger.info("Selected email is open")
        except Exception as e:
            self.logger.error(f"Error occurred:{e}")


    def click_on_attachment_option(self):
        try:
            self.click(control_title="Attachment options",control_type="Button")
        except Exception as e:
            self.logger.error(f"Failed to click on Attachment options button:{e}")


    def click_on_save_attachment(self):
        try:
            context_menu = self.get_context_menu()
            grp_box = context_menu.child_window(title=" ",control_type="Group",found_index=2)

            save_as=None

            for menu_item in grp_box.children(control_type="MenuItem"):
                if menu_item.window_text().strip().lower() == "save as".strip().lower():
                    save_as=menu_item
                    break

            if save_as:
                save_as.click_input()
                time.sleep(5)
                self.logger.info("Clicked on 'Save As' successfully.")
            else:
                self.logger.error("Failed to click on Save As")

        except Exception as e:
            self.logger.error(f"Error occurred while trying to click 'Save As': {e}")
        # self.outlook.print_control_identifiers()


    def save_attachment_at_folder(self, folder_name="pywin_practice"):
        self.save_attachment_dialog = self.outlook.child_window(title_re=".*Save Attachment.*", control_type="Window")
        self.save_attachment_dialog.wait('visible',timeout=5)
        self.logger.info("Save Attachment dialog is opened")

        tree = self.save_attachment_dialog.child_window(title="Tree View", control_type="Tree")

        target_folder = None

        for folder in tree.descendants(control_type="TreeItem"):
            if folder_name.lower() in folder.window_text().lower():
                target_folder = folder
                break

        if target_folder:
            target_folder.select()
            self.logger.info(f"Selected folder:{target_folder.window_text()}")
        else:
            self.logger.error(f"Folder {folder_name} not found.")


    def save_file_with_unique_name(self):
        file_location=r"C:\Users\swapnalik\Documents\pywin_practice"
        base_name = "Doc"

        for i in count(0):
            if i == 0:
                self.filename = f"{base_name}.txt"
            else:
                self.filename = f"{base_name}_{i}.txt"
            #Combines the folder path and the file name into a full file path.
            full_path = os.path.join(file_location, self.filename)
            #Check the file name already exists in folder if not exists break
            #Otherwise loop continues and increment by i to try the next name like Doc_1.txt, Doc_2.
            if not os.path.exists(full_path):
                break

        self.logger.info(f"File to be saved as: {self.filename}")


    def edit_file_name(self):
        edit_file_name = self.save_attachment_dialog.child_window(title="File name:", control_type="Edit")
        edit_file_name.type_keys(self.filename, with_spaces=True)
        self.logger.info(f"Edited file name: {self.filename}")


    def click_on_save(self):
        save_button = self.save_attachment_dialog.child_window(title="Save", auto_id="1", control_type="Button")
        save_button.click_input()
        self.logger.info("Save button is clicked")



