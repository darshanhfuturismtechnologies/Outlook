import time
from Helper.helper import Helper


class QuickSteps(Helper):
    def __init__(self,app):
        super().__init__(app,".*Outlook.*")
        self.mark_complete = None
        self.quick_step_dlg = None

    def click_on_create_quick_steps(self):
        try:
            quick_steps_pane = self.outlook.child_window(title="Quick Steps", control_type="Pane")
            quick_steps_pane.wait('visible', timeout=5)
            self.logger.info("Quick step pane is visible")

            list_view = quick_steps_pane.child_window(title="Quick Steps", control_type="DataGrid")

            #Use get_items() to find the list items in the DataGrid
            for list_item in list_view.get_items():
                if list_item.window_text() == "Create New":
                    list_item.click_input()
                    break
                self.logger.info("Clicked on Create New")
            else:
                self.logger.error("Failed to find  Create New")

        except Exception as e:
            self.logger.error(f"Create occured while clicking on Create New: {e}")

    def capture_quick_steps_dlg(self):
        try:
            self.quick_step_dlg = self.outlook.window(title="Edit Quick Step", control_type="Window", found_index=0)
            self.quick_step_dlg.wait('visible', timeout=10)
            self.quick_step_dlg.set_focus()
            self.logger.info("Quick Step dlg is visible")
        except Exception as e:
            self.logger.error(f"Failed to capture Quick Step dlg: {e}")

    def choose_an_action(self):
        try:
            opn=self.quick_step_dlg.child_window(title="Open", control_type="Button",found_index=0)
            opn.click_input()
            self.logger.info("Open clicked")

            action_menu=self.quick_step_dlg.child_window(title="Action", control_type="Menu",found_index=0)
            action_menu.wait('visible', timeout=5)
            self.logger.info("Action menu is visible")

            categorize_grp=action_menu.child_window(title="Categories, Tasks and Flags",control_type="Group")
            list_items = categorize_grp.descendants(control_type="ListItem")
            self.logger.info(f"Length of list_items: {len(list_items)}")

            target=None

            for item in list_items:
                if item.window_text() =="Mark complete":
                    target=item
                    break

            if target:
                target.click_input()
                self.logger.info("Mark complete selected")
            else:
                self.logger.error("Failed to find  Mark Complete")

        except Exception as e:
            self.logger.error(f"Failed to select Action combo: {e}")

    def choose_short_cut_key(self):
        try:
            #Click 'Open' button to show shortcut options
            open_shortcut_key = self.quick_step_dlg.child_window(title="Open", control_type="Button", found_index=0)
            open_shortcut_key.click_input()
            self.logger.info("Open shortcut key clicked")

            #Expand ComboBox
            shortcut_key_combo = self.quick_step_dlg.child_window(title="Shortcut key", control_type="ComboBox")
            shortcut_key_combo.expand()
            self.logger.info("Shortcut key ComboBox expanded")

            #Get all ListItems
            items = shortcut_key_combo.descendants(control_type="ListItem")
            self.logger.info(f"Found {len(items)} shortcut key items")

            #Select the target shortcut
            for item in items:
                if item.window_text() == "CTRL+SHIFT+1":
                    item.click_input()
                    self.logger.info("Shortcut key 'CTRL+SHIFT+1' selected")
                    break
            else:
                self.logger.error("Failed to find shortcut key 'CTRL+SHIFT+1'")

        except Exception as e:
            self.logger.error(f"Error in choose_short_cut_key: {e}")

    def click_on_finish_by_handling_pop_up(self):
        try:
            # After shortcut selection, a Yes/No popup may appear
            try:
                popup = self.outlook.child_window(title="Microsoft Outlook", control_type="Window")
                if popup.exists(timeout=1):
                    self.logger.info("Confirmation popup appeared")

                    # Try Yes first
                    try:
                        yes_btn = popup.child_window(title="Yes", control_type="Button")
                        if yes_btn.exists():
                            yes_btn.click_input()
                            self.logger.info("Clicked Yes on popup")
                    except:
                        pass

                    # Try No if Yes not found
                    try:
                        no_btn = popup.child_window(title="No", control_type="Button")
                        if no_btn.exists():
                            no_btn.click_input()
                            self.logger.info("Clicked No on popup")
                    except:
                        pass
            except:
                pass  # No popup → continue

            # After popup, check if dialog still exists
            if not self.quick_step_dlg.exists(timeout=1):
                self.logger.info("Quick Step dialog closed automatically after popup")
                return

            # Click Finish only if dialog is still open
            finish_btn = self.quick_step_dlg.child_window(title="Finish", control_type="Button")
            finish_btn.wait('enabled', timeout=10)
            time.sleep(1)
            finish_btn.click_input()

            self.logger.info("Finish clicked")

        except Exception as e:
            self.logger.error(f"Error in click_on_finish: {e}")

    def select_mail_mark_as_complete_on_demo_sub_mail(self):
        self.click_menu_item(".*Inbox:.*")

        try:
            # Get the container that holds all groups of emails
            table_view = self.outlook.child_window(title="Table View", control_type="Table")
            # Find all DataItem controls under that table each representing an email
            mail_items = table_view.descendants(control_type="DataItem")

            # Target Subject to find
            target_sub = "Demo subject"

            # Find first email with target subject and click it
            for email in mail_items:
                if target_sub in email.window_text():
                    email.click_input()
                    time.sleep(1)  # Wait for email visibility
                    self.logger.info(f"Clicked email with subject: {target_sub}")
                    # Email is stored as selected_mail for further use
                    self.selected_mail = email
                    # Exit the loop after clicking
                    break

            else:
                # If email is not clicked raise error
                raise Exception(f"No email found with subject: {target_sub}")
        except Exception as e:
            self.logger.error(f"Error while selecting Demo Sub Mail:{e}")

    def click_on_mark_ss_complete_inside_quick_steps(self):
        try:
            mark_complete=self.outlook.child_window(title="Mark complete",control_type="ListItem",found_index=0)
            mark_complete.click_input()
            self.logger.info("Mark complete clicked")
        except Exception as e:
            self.logger.error(f"Error in click_on_mark_complete: {e}")

    def delete_mark_complete_quick_step(self):
        try:
            self.mark_complete = self.outlook.child_window(title="Mark complete",control_type="ListItem",found_index=1)
            self.mark_complete.right_click_input()
            self.logger.info("Right clicked on Mark complete")
        except Exception as e:
            self.logger.error(f"Error while clicking right click: {e}")


    def click_on_delete(self):
        try:
            #Wait for the context menu to be visible
            context_menu = self.outlook.child_window(title="Context Menu", control_type="Menu")
            context_menu.wait('visible',timeout=5)
            self.logger.info("Context menu is visible")

            #Get the Pane inside the context menu
            pane = context_menu.child_window(control_type="Pane")
            self.logger.info("Found Pane inside context menu")

            #Find the GroupBox that contains the MenuItems
            group_box = pane.child_window(control_type="Group",found_index=0)
            self.logger.info("Found GroupBox inside Pane")

            #Get all MenuItems inside the GroupBox
            menu_items = group_box.children(control_type="MenuItem")
            self.logger.info(f"Found {len(menu_items)} menu items in the GroupBox.")

            #Search for the Delete menu item by title
            delete_btn = None
            for item in menu_items:
                if item.window_text() == "Delete":
                    delete_btn = item
                    break

            # If Delete button is found,click it
            if delete_btn:
                delete_btn.click_input()
                self.logger.info("Delete button clicked")
            else:
                self.logger.error("Delete menu item not found in the context menu.")

        except Exception as e:
            self.logger.error(f"Error in click_on_delete:{e}")























