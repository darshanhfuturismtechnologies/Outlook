import time
from Helper.helper import Helper


class Signature(Helper):
    def __init__(self, app):
        super().__init__(app,".*Outlook.*")
        self.new_sign_dlg = None
        self.sign_and_stationery_dlg = None
        self.New_email_dlg = None

    def click_on_new_mail(self):
        self.click(control_title="New Email",control_type="Button")

        self.New_email_dlg=self.wait_for_window(title_re=".*- Message..*",timeout=5)


    def click_signature_menu_select_signature_option(self):
        try:
            self.click_child_window(control_title="Signature",control_type="MenuItem")

            signature_menu=self.New_email_dlg.child_window(title="Signature",control_type="Menu")
            signature_menu.wait('visible',timeout=5)
            self.logger.info("Signature menu is visible")

            signatures_option1=signature_menu.child_window(title="Signatures...",control_type="MenuItem")
            signatures_option1.click_input()
            self.logger.info("Clicked on Signature option")

        except Exception as e:
            self.logger.error(f"Failed to click on signature_option1:{e}")

    def capture_signatures_and_stationery_window(self):
        self.sign_and_stationery_dlg=self.wait_for_child_window(parent=self.New_email_dlg,title_re=".*Signatures and Stationery.*",timeout=5)

    def click_on_new(self):
        self.click_inside_parent_win(parent=self.sign_and_stationery_dlg,title="New",control_type="Button")

    def capture_new_sign_dlg(self):
        self.new_sign_dlg=self.wait_for_child_window(parent=self.sign_and_stationery_dlg,title_re=".*New Signature.*")

    def type_name_for_this_signature(self):
        sign_name=self.new_sign_dlg.child_window(auto_id="17", control_type="Edit")
        sign_name.type_keys("Swapnali K.",with_spaces=True,pause=0.1)
        self.logger.info("Typed sign name for this signature")

        ok=self.new_sign_dlg.child_window(title="OK", control_type="Button")
        cancel=self.new_sign_dlg.child_window(title="Cancel",control_type="Button")
        close=self.new_sign_dlg.child_window(title="Close",control_type="Button")

        if ok.exists(timeout=2):
            ok.click_input()
            self.logger.info("Ok clicked")

        elif cancel.exists(timeout=2):
            cancel.click_input()
            self.logger.info("Cancel clicked")

        else:
            close.click_input()
            self.logger.info("Close clicked")

    def new_msg(self):
        try:
            combo_box = self.sign_and_stationery_dlg.child_window(title="New messages:",control_type="ComboBox")
            combo_box.wait('ready',timeout=5)
            combo_box.expand()
            self.logger.info("Expanded 'New messages' combo box")

            #Get all list items
            items = combo_box.children(control_type="ListItem")
            self.logger.info(f"Total items: {len(items)}")

            #Select your signature
            for item in items:
                if "Swapnali K." in item.window_text():
                    item.click_input()
                    self.logger.info("Selected 'Swapnali K.' successfully")
                    return

            raise Exception("Signature 'Swapnali K.' not found in New messages combo box")

        except Exception as e:
            self.logger.error(f"Failed to click on new_msg:{e}")

    def edit_signature(self):
        try:
            dlg = self.sign_and_stationery_dlg

            #The only control that can receive real keyboard input
            html = dlg.child_window(class_name="_WwG")
            html.wait("exists ready", timeout=10)

            #Focus on it
            html.click_input()
            time.sleep(0.5)

            #Clear existing signature
            html.type_keys("^a{DEL}", pause=0.05)

            #Type new signature
            html.type_keys("Best regards,{ENTER}Swapnali K.{ENTER}", with_spaces=True, pause=0.05)
            self.logger.info("Signature successfully typed in MSHTML editor")

        except Exception as e:
            self.logger.error(f"Failed to click on edit_signature:{e}")

    def click_on_ok(self):
        try:
            ok = self.sign_and_stationery_dlg.child_window(title="OK", control_type="Button")
            cancel = self.sign_and_stationery_dlg.child_window(title="Cancel", control_type="Button")

            if ok.exists(timeout=2):
                ok.click_input()
                self.logger.info("Ok clicked")

            elif cancel.exists(timeout=2):
                cancel.click_input()
                self.logger.info("Cancel clicked")

            else:
                self.logger.error("Both ok/Cancel not found")

        except Exception as e:
            self.logger.error(f"Failed to click on ok/cancel:{e}")

    def get_email_body_text(self,window):
        try:
            #MSHTML editor is usually inside a "AfxWndW" or "WebBrowser" control
            editor = window.child_window(control_type="Document")

            #Extract the text
            body_text = editor.window_text()

            self.logger.info(f"Email body captured:{body_text}")
            return body_text

        except Exception as e:
            self.logger.error(f"Failed to get email body text:{e}")
            return ""

    def close_recent_email_dlg(self):
      close_btn_for_dlg=self.New_email_dlg.child_window(title="Close",control_type="Button",found_index=0)
      close_btn_for_dlg.click_input()
      self.logger.info("Closed recent email dlg")

    def again_click_on_new_mail(self):
        self.click(control_title="New Email", control_type="Button")
        self.logger.info("New Email clicked again")

        self.New_email_dlg = self.wait_for_window(title_re=".*- Message..*",timeout=5)

        body = self.get_email_body_text(self.New_email_dlg)

        assert "Best regards" in body, "Signature not found"
        assert "Swapnali K" in body, "Signature not found"

        self.logger.info("Verified signature text in new mail window")





















