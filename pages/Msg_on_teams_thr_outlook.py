import time
from pywinauto import Application
from Helper.helper import Helper

"""Tc:-16
1.Click on New items.
2.Click on chat.
3.Integrate microsoft teams.
4.Enter recipent name and select from suggessions.
4.Edit the message and send it."""

class MsgTeams(Helper):
    def __init__(self,app):
        super().__init__(app, ".*Outlook.*")

    def click_on_new_item_menu(self):
        try:
            home_grp=self.outlook.child_window(title="Home", control_type="Group")
            self.New_item_menu=home_grp.child_window(title="New Items", auto_id="MailNewItemMenu", control_type="MenuItem")
            self.New_item_menu.wait('visible',timeout=5)
            self.New_item_menu.click_input()
            self.logger.info("New Item menu has been clicked.")
        except Exception as e:
            self.logger.error(f"Failed to click new item menu due to:{e}")

    def select_chat_option(self):
        try:
            grp_box = self.New_item_menu.child_window(control_type="Group",found_index=0)
            grp_box.wait('visible', timeout=5)

            chat_item = grp_box.child_window(title="Chat", control_type="MenuItem")

            if chat_item.exists():
                chat_item.click_input()
                self.logger.info("Chat item has been clicked.")
            else:
                self.logger.error("Chat item has been disabled.")
        except Exception as e:
            self.logger.error(f"An error occurred:{e}")

    def capture_teams_dlg(self):
        try:
            # Attach to any Teams window
            self.teams_app = Application(backend="uia").connect(title_re=".*New Message.*",timeout=10)
            self.teams_dlg = self.teams_app.window(title_re=".*New Message.*|.*Swapnali K.*")

            self.teams_dlg.wait('visible',timeout=15)
            self.logger.info("Teams dialog captured successfully.")
            return self.teams_dlg

        except Exception as e:
            self.logger.error(f"Failed to capture Teams dialog: {str(e)}")
            raise

    def enter_field_to(self):
        try:
            pane=self.teams_dlg.child_window(title="New Message | Microsoft Teams", auto_id="RootWebArea", control_type="Document")
            pane.wait('visible', timeout=5)
            self.logger.info("Pane is visible")

            to_field_comboBox = pane.child_window(title="To: ", control_type="ComboBox")
            to_field_comboBox.wait('enabled',timeout=5)
            to_field_comboBox.click_input()
            self.logger.info("To field combobox has been clicked.")

            edit_to_field = to_field_comboBox.child_window(auto_id="people-picker-input", control_type="Edit")
            edit_to_field.wait('enabled',timeout=5)
            edit_to_field.click_input()

            # Type the recipient slowly
            edit_to_field.type_keys("Swapnali K", with_spaces=True, pause=0.2)
            self.logger.info("Name typed into 'To' field.")

            #Wait for the suggestion list to appear
            suggestions_list = self.teams_dlg.child_window(control_type="List")
            suggestions_list.wait('visible', timeout=10)
            self.logger.info("Suggestions list is visible.")

            #Select the first suggestion
            first_suggestion = suggestions_list.child_window(control_type="ListItem",found_index=0)
            first_suggestion.click_input()
            self.logger.info("First suggestion clicked.")

        except Exception as e:
            self.logger.error(f"Failed to enter recipient in 'To' field:{e}")
            raise

    def edit_and_send_msg(self):
        try:
            Swapnali_dlg = self.teams_app.window(title_re=".*Swapnali K.*")
            Swapnali_dlg.wait('visible', timeout=10)
            self.logger.info("Swapnali K's window is visble")

            edit_msg = self.teams_dlg.child_window(title="Type a message\n", control_type="Edit")
            edit_msg.type_keys("Hellooo,How are you..!", with_spaces=True, pause=0.1)
            self.logger.info("Message typed into 'Edit' field.")
        except Exception as e:
            self.logger.error(f"Failed to edit message: {str(e)}")

    def send_msg(self):
        try:
            send_btn = self.teams_dlg.child_window(title="Send (Ctrl+Enter)", control_type="Button")
            send_btn.wait('enabled', timeout=5)
            send_btn.click_input()
            self.logger.info("Send message has been clicked.")
        except Exception as e:
            self.logger.error(f"Failed to send message: {str(e)}")

    def close_microsoft_teams_app(self):
        try:
            self.teams_dlg.close()
            self.logger.info("Microsoft Teams app has been closed.")
        except Exception as e:
            self.logger.error(f"Failed to close microsoft teams: {str(e)}")




































