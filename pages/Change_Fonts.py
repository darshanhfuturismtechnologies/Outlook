import time
from Helper.helper import Helper


class ChangeFont(Helper):
    def __init__(self, app):
        super().__init__(app,".*Outlook.*")
        self.font_dlg = None
        self.sign_and_stationery_dlg = None
        self.option_dlg = None

    def click_on_file(self):
        self.click(control_title="File Tab",control_type="Button")
        
        try:
            file_list = self.outlook.child_window(title="File", auto_id="NavBarMenu", control_type="List")
            options = file_list.child_window(title="Options", control_type="ListItem")
            options.click_input()
            self.logger.info("Clicked on options")
        except Exception as e:
            self.logger.error(f"Failed to click on options: {e}")

        
    def click_on_mail(self):
        try:
            self.option_dlg = self.outlook.window(title_re=".*Outlook Options.*", control_type="Window")
            self.option_dlg.wait("visible", timeout=5)
            self.logger.info("Option dlg is visible")

            option_grp = self.option_dlg.child_window(title="Outlook Options", control_type="Group")
            mail = option_grp.child_window(title="Mail", control_type="ListItem")
            mail.click_input()
            self.logger.info("Clicked on mail")
            time.sleep(1)                                                                              #Time to change the UI
        except Exception as e:
            self.logger.error(f"Failed to click on mail:{e}")


    def click_on_stationery_and_font(self):
        self.click_inside_parent_win(parent=self.option_dlg,title="Stationery and Fonts...", control_type="Button")
        time.sleep(5)

    def capture_signature_and_stationery_dlg(self):
        self.sign_and_stationery_dlg=self.wait_for_child_window(parent=self.option_dlg,title_re=".*Signatures and Stationery.*",timeout=5)

    def select_first_font(self):
        try:
            font_btn = self.sign_and_stationery_dlg.child_window(title="Font...", control_type="Button", found_index=0)
            font_btn.click_input()
            self.logger.info("Selected font btn")

            self.font_dlg = self.sign_and_stationery_dlg.child_window(title="Font", control_type="Window")
            self.font_dlg.wait("exists enabled visible ready", timeout=5)
            self.logger.info("Font dlg is visible")
            self.font_dlg.print_control_identifiers()
        except Exception as e:
            self.logger.error(f"Failed to click on font btn:{e}")

    def edit_font(self):
        aptos=self.font_dlg.child_window(title="Aptos", control_type="ListItem")
        aptos.click_input()
        self.logger.info("Clicked on Aptos")


















        









