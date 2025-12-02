from logging import exception
from Helper.helper import Helper


class Draft_mail(Helper):
    def __init__(self, app,test_data):
        super().__init__(app,".*Outlook.*")
        self.test_data=test_data
        self.new_email_dlg = None

    def click_on_new_email(self):
        self.click(control_title="New Email",control_type="Button")

    def capture_new_email_dlg(self):
        self.new_email_dlg=self.wait_for_window(title_re=".*- Message..*",control_type="Window")
        
    def enter_recipient(self):
        try:
            to = self.new_email_dlg.child_window(auto_id="4117", control_type="Edit")
            to.type_keys(self.test_data["to"]+"{ENTER}", pause=0.1)
            self.logger.info(F"Recipient name {self.test_data["to"]} has been entered")
        except exception as e:
            self.logger.error(f"Failed to enter recipient: {e}")


    def enter_cc(self):
        try:
            cc = self.new_email_dlg.child_window(auto_id="4126", control_type="Edit")
            cc.type_keys(self.test_data["cc"]+"{ENTER}", pause=0.1)
            self.logger.info(f"CC {self.test_data["cc"]} has been entered")
        except exception as e:
            self.logger.error(f"Failed to enter cc: {e}")

    def enter_subject(self):
        try:
            subject=self.new_email_dlg.child_window(auto_id="4101", control_type="Edit")
            subject.type_keys(self.test_data["subject"]+"{ENTER}",pause=0.1,with_spaces=True)
            self.logger.info(f"Subject {self.test_data['subject']} has been entered")
        except exception as e:

            self.logger.error(f"Failed to enter subject: {e}")
    def enter_body(self):
        try:
            # splits this string into a list of lines wherever a newline character \n occurs.
            edit_body=self.new_email_dlg.child_window(title="Page 1 content", auto_id="Body", control_type="Edit")
            lines=self.test_data["body"].split("\n")
            for line in lines:
                edit_body.type_keys(line, pause=0.1, with_spaces=True)
                edit_body.type_keys("{ENTER 2}")              #Move to next line
            self.logger.info("Text in body is set correctly")
        except Exception as e:
            self.logger.error(f"Failed to enter body: {e}")

    def click_on_close(self):
        close_icon=self.new_email_dlg.child_window(title="Close",control_type="Button",found_index=0)
        close_icon.click_input()

    def handle_popup_and_click_on_ok(self):
        try:
            microsoft_popup = self.new_email_dlg.child_window(title="Microsoft Outlook", control_type="Window")
            microsoft_popup.wait('visible', timeout=5)
            self.logger.info("Microsoft popup has been appeared")

            yes = microsoft_popup.child_window(title="Yes", auto_id="6", control_type="Button")
            no = microsoft_popup.child_window(title="No", auto_id="7", control_type="Button")
            cancel = microsoft_popup.child_window(title="Cancel", auto_id="2", control_type="Button")

            if yes.exists(timeout=2):
                yes.click_input()
                self.logger.info("Yes clicked")

            elif no.exists(timeout=2):
                no.click_input()
                self.logger.info("No clicked")

            elif cancel.exists(timeout=2):
                cancel.click_input()
                self.logger.info("Cancel clicked")

            else:
                raise RuntimeError("No expected button (Yes/No/Cancel) found on popup")

        except Exception as e:
            self.logger.error(f"Failed to handle popup: {e}")

    def click_on_draft(self):
        self.click_menu_item(".*Draft.*")

        table_view=self.outlook.child_window(title="Table View", auto_id="4704", control_type="Table")
        mail_items=table_view.descendants(control_type="DataItem")
        self.logger.info(f"Found {len(mail_items)} mail items")

        if mail_items:
            # Assuming the first mail is the most recent
            recent_mail = mail_items[0]
            print(f"Most recent email: {recent_mail}")
            recent_mail.click_input()
            self.selected_mail = recent_mail
            self.logger.info(f"Most recent email: {recent_mail}")
        else:
            raise Exception("No emails found in the inbox")

        #Verify draft mail appeared in draft or not by assertion
        #Draft mail is not found yet
        draft_found = False

        #This loop itterate one by one through each mail
        for item in mail_items:
            #Verify the expected subject is present in draft mail
            if self.test_data["subject"] in item.window_text():
                #If subject matches we mark as yes/true
                draft_found = True
                #Once draft found stop loop immediately
                self.logger.info(f"Draft mail found with subject: {self.test_data['subject']}")
                break

        if not draft_found:
            error_msg = (f"Draft mail with subject '{self.test_data['subject']}' "
                f"NOT found in Draft folder"
            )
            self.logger.error(error_msg)
            raise AssertionError(error_msg)




