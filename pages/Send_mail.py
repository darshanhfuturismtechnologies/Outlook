import time
from Helper.helper import Helper


"""TC_03:-
1.Click on inbox,
2.Enter details like To,CC,Subject,
3.Enter text in msg body
4.Click on attach file and attach file as a copy
5.Click on Send"""


class SendMail(Helper):
    def __init__(self,app):
        super().__init__(app,".*Outlook.*")
        self.send_mail_window = None


    def click_on_send_mail(self):
        try:
            self.click(control_title="New Email", control_type="Button")
            self.send_mail_window = self.app.window(title_re=".*- Message..*")
            self.send_mail_window.wait('visible', timeout=10)
            self.logger.info("Clicked on New mail and Send mail window opened")
        except Exception as e:
            self.logger.error(f"Failed to click on send mail:{e}")
            raise


    def enter_to_mail_id(self):
        try:
            To_field = self.send_mail_window.child_window(auto_id="4117",control_type="Edit")
            To_field.type_keys("swapnalik@futurismtechnologies.com",pause=0.1,with_spaces=True)
            To_field.type_keys('{ENTER}')
            self.logger.info("To email id is set correctly")
        except Exception as e:
            self.logger.error(f"Failed to enter to mail id:{e}")
            raise


    def enter_cc_mail_id(self):
        try:
            CC_field = self.send_mail_window.child_window(auto_id="4126", control_type="Edit")
            CC_field.type_keys("swapnalik@futurismtechnologies.com",pause=0.1,with_spaces=True)
            CC_field.type_keys('{ENTER}')
            self.logger.info("Cc email id is set correctly")
        except Exception as e:
            self.logger.error(f"Failed to enter to cc mail id:{e}")
            raise


    def enter_subject(self):
        try:
            subject = self.send_mail_window.child_window(auto_id="4101", control_type="Edit")
            subject.type_keys("Demo subject", pause=0.1, with_spaces=True)
            subject.type_keys('{ENTER}')
            self.logger.info("Subject is set correctly")
        except Exception as e:
            self.logger.error(f"Failed to enter subject:{e}")
            raise


    def enter_text_in_body(self):
        try:
            body = self.send_mail_window.child_window(title="Page 1 content", auto_id="Body", control_type="Edit")
            body.type_keys("Hello, Swapnali K.", pause=0.1, with_spaces=True)
            body.type_keys('{ENTER 2}')
            body.type_keys("This is the body of the page....!", pause=0.1, with_spaces=True)
            body.type_keys('{ENTER 2}')
            body.type_keys("Best Regards,", pause=0.1, with_spaces=True)
            body.type_keys('{ENTER}')
            body.type_keys("Swapnali K.", pause=0.1, with_spaces=True)
            self.logger.info("Text in body is set correctly")
        except Exception as e:
            self.logger.error(f"Failed to enter body text:{e}")
            raise


    def attach_file_to_mail(self):
        self.click_child_window(control_title="Attach File...",control_type="MenuItem")
        recent_items =self.send_mail_window.window(title_re=".*Recent Items.*")
        recent_items.wait('visible',timeout=10)
        self.logger.info("Recent items window is open")

        #file name that you want to attach
        target_file_path = "Copy of QA_intern_training_desktop_automation (002).xlsx"
        #This calls a helper method to find and click the list item that matches the file name.
        self.click_list_item_by_text(recent_items,target_file_path)

    def attach_file_as_copy(self):
        try:
            attach_file_popup = self.send_mail_window.child_window(title_re=".*How do you want to attach this file?.*",control_type="Window",found_index=0)
            attach_file_popup.wait('ready',timeout=20)
            self.logger.info("Attach Popup is occurred..")

            attach_file_popup.child_window(title="Attach as copy", control_type="Button").click_input()
            time.sleep(5)
            self.logger.info("Attach as Copy option is selected")
        except Exception as e:
            self.logger.error(f"Failed to attach file: {e}")


    def click_on_send(self):
        self.click_child_window(control_title="Send", control_type="Button")
        #Email takes 30-40 seconds to sent and then it close window
        self.wait_for_window_to_close(self.send_mail_window,timeout=50)













































     # try:
        #     attach_file_popup = self.send_mail_window.child_window(title_re=".*How do you want to attach this file?.*",control_type="Window",found_index=0)
        #     attach_file_popup.wait('ready',timeout=20)
        #     self.logger.info("Attach Popup is occurred..")
        #
        #     attach_file_popup.child_window(title="Attach as copy", control_type="Button").click_input()
        #     time.sleep(5)
        #     self.logger.info("Attach as Copy option is selected")
        # except Exception as e:
        #     self.logger.error(f"Failed to attach file: {e}")
        #     raise































