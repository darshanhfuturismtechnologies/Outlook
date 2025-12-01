import time
from pywinauto.keyboard import send_keys
from Helper.helper import Helper

"""Tc_11:
1.Click on search bar,search for contact and select contact.
2.Click on create new group.
3.Enter group name.
4.Click on add members and select members from address book.
5.Click on ok.
6.Click on save and close.
7.Click on New Item menu and select Email Message.
8.Enter group name in to field,enter sub,enter msg body and click on send."""

class GroupMail(Helper):
    def __init__(self,app,test_data):
        super().__init__(app,".*Outlook.*")
        self.test_data = test_data
        self.send_mail_window = None
        self.contact_window = None
        self.select_member_dlg = None
        self.new_contact_grp_dlg = None

    #Click on search bar,search for contact and select contact.
    def click_on_search_bar_and_enter_contact(self):
        try:
            tell_me_box = self.outlook.child_window(auto_id="TellMeTextBoxAutomationId", control_type="Edit").wait('visible', timeout=5)
            tell_me_box.type_keys(self.test_data["contacts"])
            self.outlook.child_window(title="Contacts",control_type="MenuItem").click_input()
            self.logger.info("Clicked on contact")
        except Exception as e:
            self.logger.error(f"Clicked on contact via search box:{e}")

    #Click on create new group.
    def click_on_new_group(self):
        try:
            self.contact_window = self.app.window(title_re=".*Contacts.*")
            self.contact_window.wait('exists visible', timeout=10)

            self.click(control_title="New Contact Group", control_type="Button")

            self.new_contact_grp_dlg = self.app.window(title_re=".* - Contact Group.*")
            self.new_contact_grp_dlg.wait('visible', timeout=10)
            self.logger.info("New contact grp dialog opened")
        except Exception as e:
            self.logger.error(f"Failed to open new contact group:{e}")

    #Enter group name.
    def enter_grp_name(self):
        try:
            name_edit = self.new_contact_grp_dlg.child_window(auto_id="4096",control_type="Edit").wait('visible',timeout=5)
            name_edit.set_edit_text(self.test_data["group_name"])
            self.logger.info(f"Entered contact group name:{self.test_data["group_name"]}")
        except Exception as e:
            self.logger.error(f"Failed to enter contact group name:{e}")


    #Click on Add Members menu
    def click_on_add_members(self):
        try:
            add_members_btn = self.new_contact_grp_dlg.child_window(title="Add Members",auto_id="DistributionListNewMemberMenu",control_type="MenuItem")
            add_members_btn.click_input()
            self.logger.info("Clicked on Add Members button")

            self.new_contact_grp_dlg.child_window(title="From Address Book", control_type="MenuItem").click_input()
            self.logger.info("Clicked on From Address Book button")
        except Exception as e:
            self.logger.error(f"Failed to click on Address Book:{e}")

    #Select members from Address book
    def select_members(self):
        try:
            self.select_member_dlg = self.new_contact_grp_dlg.child_window(
                title_re=".*Select Members: Global Address List.*", control_type="Window")
            self.logger.info("Select member dialog is opened")

            # Enter first name in search box and press enter
            search_box = self.select_member_dlg.child_window(auto_id="101", control_type="Edit").wait('visible',timeout=5)
            search_box.click_input()
            search_box.type_keys(self.test_data["member1"], with_spaces=True, pause=0.1)
            self.logger.info(f"Entered:{self.test_data["member1"]}and pressed Enter")
            time.sleep(1)  # small delay for searching name
            send_keys("{ENTER}")
            time.sleep(1)
            # Clear the search bar
            search_box.click_input()
            # Clear previous name
            send_keys('^a{BACKSPACE}')
            self.logger.info("Search bar is cleared")
            # Enter second name in search box and press enter
            search_box.type_keys(self.test_data["member2"]+"{ENTER}", with_spaces=True, pause=0.1)
            self.logger.info(f"Entered:{self.test_data["member2"]} and pressed Enter")
            time.sleep(1)
        except Exception as e:
            self.logger.error(f"Failed to Add members:{e}")
            raise

    #Click on Ok/Cancel/Close icon
    def click_on_ok_or_cancel(self):
        try:
            ok_btn = self.select_member_dlg.child_window(title="OK", auto_id="100", control_type="Button")
            ok_btn.click_input()
            self.logger.info("Clicked on OK button")
        except Exception as e:
            self.logger.error(f"Failed to click OK button: {e}")
            try:
                cancel_btn = self.new_contact_grp_dlg.child_window(title="Cancel", auto_id="2", control_type="Button")
                cancel_btn.click_input()
                self.logger.info("Clicked on Cancel button")
            except Exception as e:
                self.logger.error(f"Failed to click Cancel button: {e}")
                try:
                    close_icon = self.new_contact_grp_dlg.child_window(title="Close", control_type="Button")
                    close_icon.click_input()
                    self.logger.info("Clicked on Close button")
                except Exception as e:
                    self.logger.error(f"Failed to click Close button: {e}")
            time.sleep(3)

    #Click on save and close menu
    def click_on_save_and_close(self):
        try:
            save_and_close=self.new_contact_grp_dlg.child_window(title="Save & Close", auto_id="SaveAndClose", control_type="Button")
            save_and_close.wait('visible', timeout=10)
            save_and_close.click_input()
            self.logger.info("Clicked on Save & Close button")
        except Exception as e:
            self.logger.error(f"Failed to click save & close button: {e}")

    #Click on New Item Menu and select Email message
    def click_on_new_item(self):
        self.contact_window.wait('exists visible', timeout=10)
        try:
            # Click on New Items menu
            new_item = self.contact_window.child_window(title_re="New Items", control_type="MenuItem")
            new_item.wait('visible enabled', timeout=10)
            new_item.click_input()
            self.logger.info("Clicked on New Item button")

            #Get all menu items descendants
            sub_menus = self.contact_window.descendants(control_type="MenuItem")

            #Loop through and click Email msg
            for item in sub_menus:
                if item.window_text() == "E-mail Message":
                    item.click_input()
                    self.logger.info("Clicked on E-mail Message menu item")
                    break                                                    #Break stops the loop  when correct menu is found.
            else:
                self.logger.error("E-mail Message menu item not found")

        except Exception as e:
            self.logger.error(f"Failed to click on New Item: {e}")
        time.sleep(3)

    def enter_group_name_in_to_field(self):
        try:
            # Wait until the mail window appear
            self.send_mail_window = self.app.window(title_re=".*- Message..*")
            self.send_mail_window.wait('visible', timeout=10)
            self.logger.info("Send mail window is opened")

            # Enter Group name in To field
            To_field = self.send_mail_window.child_window(auto_id="4117", control_type="Edit")
            To_field.type_keys(self.test_data["group_name"], pause=0.1, with_spaces=True)
            To_field.type_keys('{ENTER}')
            self.logger.info(f"Group name:{self.test_data["group_name"]} is entered")
            self.logger.info("To field is entered")
        except Exception as e:
            self.logger.error(f"Failed to enter group name: {e}")

    #Enter subject
    def enter_subject(self):
        try:
            subject = self.send_mail_window.child_window(auto_id="4101", control_type="Edit")
            subject.type_keys(self.test_data["subject"], pause=0.1, with_spaces=True)
            subject.type_keys('{ENTER}')
            self.logger.info(f"Subject is:{self.test_data["subject"]} entered")
        except Exception as e:
            self.logger.error(f"Failed to enter subject: {e}")

    #Enter text in body
    def enter_text_in_body(self):
        try:
            body = self.send_mail_window.child_window(title="Page 1 content", auto_id="Body", control_type="Edit")
            lines = self.test_data["body"].split("\n")
            for line in lines:
                body.type_keys(line, pause=0.1, with_spaces=True)
                body.type_keys("{ENTER2}")
        except Exception as e:
            self.logger.error(f"Failed to enter text:{e}")

    #Click on Send
    def click_on_send(self):
        self.click_child_window(control_title="Send", control_type="Button")
        #Email takes 30-40 seconds to sent and then it close window
        self.wait_for_window_to_close(self.send_mail_window,timeout=60)


























