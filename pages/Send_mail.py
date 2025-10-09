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

    def click_on_send_mail(self):
        self.click(control_title="New Email", control_type="Button")
        self.send_mail_window = self.app.window(title_re=".*- Message..*")
        self.send_mail_window.wait('visible', timeout=10)

    def enter_to_mail_id(self):
        To_field = self.send_mail_window.child_window(auto_id="4117", control_type="Edit")
        To_field.type_keys("swapnalik@futurismtechnologies.com",pause=0.1,with_spaces=True)
        To_field.type_keys('{ENTER}')

    def enter_cc_mail_id(self):
        CC_field=self.send_mail_window.child_window(auto_id="4126", control_type="Edit")
        CC_field.type_keys("swapnalik@futurismtechnologies.com",pause=0.1,with_spaces=True)
        CC_field.type_keys('{ENTER}')

    def enter_subject(self):
        subject=self.send_mail_window.child_window(auto_id="4101", control_type="Edit")
        subject.type_keys("Demo subject",pause=0.1,with_spaces=True)
        subject.type_keys('{ENTER}')

    def enter_text_in_body(self):
        body=self.send_mail_window.child_window(title="Page 1 content", auto_id="Body", control_type="Edit")
        body.type_keys("Hello, Swapnali K.",pause=0.1,with_spaces=True)
        body.type_keys('{ENTER 2}')
        body.type_keys("This is the body of the page....!",pause=0.1,with_spaces=True)
        body.type_keys('{ENTER 2}')
        body.type_keys("Best Regards,",pause=0.1,with_spaces=True)
        body.type_keys('{ENTER}')
        body.type_keys("Swapnali K.",pause=0.1,with_spaces=True)

    def attach_file_to_mail(self):
        self.click_child_window(control_title="Attach File...", control_type="MenuItem")
        recent_items = self.send_mail_window.window(title_re=".*Recent Items.*")
        recent_items.wait('visible', timeout=10)

        #file name that you want to attach
        target_file_path = "Copy of QA_intern_training_desktop_automation (002).xlsx"
        #This calls a helper method to find and click the list item that matches the file name.
        self.click_list_item_by_text(recent_items,target_file_path)

        try:
            attach_file_popup = self.send_mail_window.child_window(title_re=".*How do you want to attach this file?.*",control_type="Window", found_index=0)
            attach_file_popup.wait('visible', timeout=20)
            attach_file_popup.child_window(title="Attach as copy", control_type="Button").click_input()
            time.sleep(2)
        except Exception as e:
            self.logger.error(e)

    def click_on_send(self):
        self.click_child_window(control_title="Send", control_type="Button")
        #Email takes 30-40 seconds to sent and then it close window
        self.wait_for_window_to_close(self.send_mail_window,timeout=60)

































