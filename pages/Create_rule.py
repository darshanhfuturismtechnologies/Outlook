import time

from Helper.helper import Helper


class Create_rule(Helper):
    def __init__(self,app):
        super().__init__(app, ".*Outlook.*")
        self.rules_and_alert_dlg = None
        self.edit_list = None
        self.rules_wizard_dlg = None

    def navigate_to_file_menu(self):
        try:
            tool_bar = self.outlook.child_window(control_type="ToolBar", class_name="MsoCommandBar", found_index=0)
            tool_bar.child_window(title="File Tab", auto_id="FileTabButton").click_input()
            self.logger.info("File tab clicked")
        except Exception as e:
            self.logger.error(f"Failed to navigate to file tab: {e}")

    def click_on_rule_and_alert(self):
        try:
            backstage_view = self.outlook.child_window(title="Backstage view",auto_id="BackstageView",control_type="Pane")
            backstage_view.wait("exists visible",timeout=5)

            info_group = backstage_view.child_window(title="Info",control_type="Group")
            info_group.wait("ready",timeout=5)

            rules_group = info_group.child_window(title="Rules and Alerts", auto_id="GroupRulesAndAlerts",control_type="Group")
            rules_group.wait("ready",timeout=5)

            manage_btn = rules_group.child_window(title="Manage Rules & Alerts...", control_type="Button")
            manage_btn.wait("enabled visible ready", timeout=5)
            manage_btn.click_input()
            self.logger.info("Clicked on 'Manage Rules & Alerts successfully.")

        except Exception as e:
            self.logger.error(f" Failed to click 'Manage Rules & Alerts': {e}")
            raise

    def rule_and_alert_window(self):
        try:
            self.rules_and_alert_dlg = self.outlook.child_window(title="Rules and Alerts", control_type="Window")
            self.rules_and_alert_dlg.wait("visible", timeout=5)
            self.logger.info("Rule and Alert window is visible")
        except Exception as e:
            self.logger.error(f"Failed to open Rules and Alert window: {e}")

    def click_on_create_rule(self):
        self.click(control_title="New Rule...", control_type="Button")

    def handle_rules_wizard_window(self):
        try:
            self.rules_wizard_dlg=self.outlook.child_window(title="Rules Wizard", control_type="Window")
            self.rules_wizard_dlg.wait("visible", timeout=5)
            self.logger.info("Rules Wizard window is visible")
        except Exception as e:
            self.logger.error(f"Failed to open Rules Wizard window: {e}")

        try:
            next_btn1 = self.rules_wizard_dlg.child_window(title="Next >",control_type="Button")
            next_btn1.wait("ready",timeout=5)
            next_btn1.click_input()
            self.logger.info("Next > in Wizard window clicked")
        except Exception as e:
            self.logger.error(f"Failed to click Next > in Wizard window:{e}")

    def set_conditions_inside_list(self):
        try:
            list_box = self.rules_wizard_dlg.child_window(title="Step 1: Select condition(s)", control_type="List",found_index=1)

            #Condition1=Uncheck from people or public group
            try:
                condition1 = list_box.child_window(title="from people or public group", control_type="CheckBox")
                if condition1.get_toggle_state():  # If already checked
                    condition1.toggle()  # Uncheck it
                    self.logger.info("Condition1:Unchecked from people or public group")
                else:
                    self.logger.info("Condition1 was already unchecked")
            except Exception as e:
                self.logger.warning(f"Condition1 not found:{e}")


            #Condition2=Checked with specific words in the subject
            try:
                condition2 = list_box.child_window(title="with specific words in the subject", control_type="CheckBox"
                )
                if not condition2.get_toggle_state():
                    condition2.toggle()
                self.logger.info("Condition2:Checked with specific words in the subject)")
            except Exception as e:
                self.logger.warning(f"Condition2 not found:{e}")

            #Condition 3:Checked through the specified account.
            try:
                condition3 = list_box.child_window(title="through the specified account", control_type="CheckBox")
                if not condition3.get_toggle_state():
                    condition3.toggle()
                self.logger.info("Condition3:Checked through the specified account")
            except Exception as e:
                self.logger.warning(f"Condition3 not found:{e}")

            #Condition 4:Checked sent only to me
            try:
                condition4 = list_box.child_window(title="sent only to me", control_type="CheckBox")
                if not condition4.get_toggle_state():
                    condition4.toggle()
                self.logger.info("Condition4: Checked sent only to me")
            except Exception as e:
                self.logger.warning(f"Condition4 not found:{e}")

        except Exception as e:
            self.logger.error(f"Failed to set conditions: {e}")

    def edit_rule_one(self):
        self.edit_list = self.rules_wizard_dlg.child_window(title="Step 2: Edit the rule description (move with arrows and press SPACEBAR to edit parameters)",control_type="List",found_index=1)
        rule1 = self.edit_list.child_window(title="through the specified account", control_type="ListItem")

        rule1.type_keys("{SPACE}")
        self.logger.info("Clicked via keystrokes")

        #Wait for the dialog
        account_dlg=self.rules_wizard_dlg.child_window(title_re=".*Account.*",control_type="Window")

        if account_dlg.wait("visible",timeout=10):
            self.logger.info("Dialog for through the specified account appeared.")
            ok_btn1 = account_dlg.child_window(title="OK", auto_id="1", control_type="Button")
            ok_btn1.click_input()
            self.logger.info("Clicked OK button in Account dialog")
        else:
            self.logger.error("Account dialog not appear")

    def edit_rule_two(self):
        #Set the focus for Rule wizard window bcause after clicking an "ok" it lost focus
        self.rules_wizard_dlg.set_focus()
        self.edit_list = self.rules_wizard_dlg.child_window(title="Step 2: Edit the rule description (move with arrows and press SPACEBAR to edit parameters)",control_type="List",found_index=1)
        rule2 = self.edit_list.child_window(title="with specific words in the subject", control_type="ListItem")

        rule2.type_keys("{SPACE}")
        self.logger.info("Clicked 'with specific words in the subject' via keystrokes")

        #Wait for dialog
        search_text_dlg=self.rules_wizard_dlg.child_window(title="Search Text",control_type="Window")

        if search_text_dlg.wait("visible",timeout=10):
            self.logger.info("Search Text dlg is visible")
            text_box=search_text_dlg.child_window(title=" ", auto_id="6908", control_type="Edit")
            text_box.type_keys("Demo Subject",with_spaces=True,pause=0.1)
            self.logger.info("Text is entered in search box")

            #Click on Add subject
            Add_btn=search_text_dlg.child_window(title="Add", auto_id="6910", control_type="Button")
            Add_btn.click_input()
            self.logger.info("Clicked on Add button")

            #Click on OK
            Ok_btn2=search_text_dlg.child_window(title="OK", auto_id="1", control_type="Button")
            Ok_btn2.click_input()
            self.logger.info("Clicked on OK button")

    def edit_rule_three(self):
        self.rules_wizard_dlg.set_focus()
        self.edit_list = self.rules_wizard_dlg.child_window(title="Step 2: Edit the rule description (move with arrows and press SPACEBAR to edit parameters)",control_type="List", found_index=1)
        rule3 = self.edit_list.child_window(title="move it to the specified folder", control_type="ListItem")

        rule3.type_keys("{SPACE}")
        self.logger.info("Clicked 'with specific words in the subject' via keystrokes")

        #Wait for dlg
        rules_and_alert_dlg=self.rules_wizard_dlg.child_window(title="Rules and Alerts", control_type="Window")

        if rules_and_alert_dlg.wait("visible",timeout=10):
            self.logger.info("Rules and Alerts dlg is visible")

            folder_tree = rules_and_alert_dlg.child_window(title="Choose a folder:", auto_id="4513",control_type="Tree")
            all_items = folder_tree.descendants(control_type="TreeItem")

            target_folder = None

            for item in all_items:
                if "Inbox" in item.window_text():
                    target_folder = item
                    break

            if target_folder:
                target_folder.click_input()
                self.logger.info("Clicked on Inbox folder")
            else:
                self.logger.error("Inbox folder not found")

            #Click OK
            ok_btn = rules_and_alert_dlg.child_window(title="OK", auto_id="1", control_type="Button")
            ok_btn.click_input()
            self.logger.info("Clicked OK button")

            #Click on Next
            next_btn2=self.rules_wizard_dlg.child_window(title="Next >", control_type="Button")
            next_btn2.wait("ready", timeout=5)
            next_btn2.click_input()
            self.logger.info("Clicked Next button 2")

            #Unchech stop processing more rules.
    def uncheck_process_rule_and_click_next(self):
            try:
                process_rule = self.rules_wizard_dlg.child_window(title="stop processing more rules",control_type="CheckBox",found_index=1)
                if process_rule.get_toggle_state():  # If already checked
                    process_rule.toggle()  #Uncheck it
                    self.logger.info("process_rule:Unchecked from people or public group")
                else:
                    self.logger.info("process_rule was already unchecked")
            except Exception as e:
                self.logger.warning(f"process_rule not found:{e}")

            next_btn3 = self.rules_wizard_dlg.child_window(title="Next >", control_type="Button")
            next_btn3.wait("ready", timeout=5)
            next_btn3.click_input()
            self.logger.info("Clicked Next button 3")

    def click_on_next_btn_four(self):
            next_btn4=self.rules_wizard_dlg.child_window(title="Next >", control_type="Button")
            next_btn4.wait("ready", timeout=5)
            next_btn4.click_input()
            self.logger.info("Clicked Next button 4")

    def click_on_finish_btn(self):
            finish_btn=self.rules_wizard_dlg.child_window(title="Finish", control_type="Button")
            finish_btn.click_input()
            self.logger.info("Clicked Finish button")

    def handle_microsoft_popup(self):
            microsoft_dlg=self.rules_wizard_dlg.child_window(title="Microsoft Outlook", control_type="Window")
            microsoft_dlg.wait("visible",timeout=5)
            click_ok=microsoft_dlg.child_window(title="OK", control_type="Button")
            click_ok.click_input()
            self.logger.info("Clicked OK button")

    def apply_rule_and_click_ok(self):
            apply=self.rules_and_alert_dlg.child_window(title="Apply", control_type="Button")
            apply.click_input()
            self.logger.info("Clicked Apply button")

            confirm_ok=self.rules_and_alert_dlg.child_window(title="OK", control_type="Button")
            confirm_ok.wait("ready exists", timeout=5)
            confirm_ok.click_input()
            self.logger.info("Clicked OK button")






























