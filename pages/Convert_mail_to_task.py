import time
from pywinauto.keyboard import send_keys
from Helper.helper import Helper

"""Tc:12
1.Click on inbox and Select the email of demo subject
2.Right click and select move<<other folder<<Task<<ok
3.A new task window opens with the email content in the description.
4.Set Due date, Priority, and Reminder,reminder date and time.
5.Save & Close.
6.Verify the The task now appears in your Tasks list and can also show in your To-Do Bar if you have it enabled.
7.Right click on task and select delete.
"""

class TaskConverter(Helper):
    def __init__(self,app):
        super().__init__(app,".*Outlook.*")
        self.task_window = None
        self.outlook.set_focus()    #Take main Outlook window to front
        self.selected_mail = None
        # self.outlook.print_control_identifiers()

    def click_on_deleted_items(self):
        self.click_menu_item(".*Sent Items:.*")

    def click_on_demo_sub_mail(self):
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

    #Right click on mail
    def right_click_on_mail(self):
        self.right_click_on_selected_option()

    #Click on move
    def click_on_move(self):
        try:
            #Wait for the right-click context menu to appear
            context_menu = self.outlook.child_window(title="Context Menu",control_type="Menu")
            context_menu.wait('visible',timeout=10)

            #Click Move menu item
            move = context_menu.child_window(title="Move", control_type="MenuItem", found_index=0)
            move.click_input()
            #Wait for submenu to appear
            time.sleep(0.5)
            self.logger.info("Move clicked")

            #Wait for the Move submenus to appear
            move_submenus = self.outlook.child_window(title="Move", control_type="Menu", found_index=0)
            move_submenus.wait('visible', timeout=5)

            #Loop through each item in the Move submenu children to find Other Folder
            for item in move_submenus.descendants(control_type="MenuItem"):
                #Check if the item text contains other folder
                if "other folder" in item.window_text().lower():
                    item.click_input()
                    self.logger.info("Other Folder clicked")
                    #Exit the loop after clicking menu
                    break
                else:
                    self.logger.error("'Other Folder...'not found in Move submenu")

        except Exception as e:
            #Catch unexpected exception
            self.logger.error(f"Failed to click 'Move->Other Folder':{e}")
            #if an error happens, throw/re-raise the exception again
            raise

    #Click task submenu
    def move_item_to_task(self):
        try:
            # Wait for the Move Items dialog to appear
            move_dialog = self.outlook.child_window(title="Move Items", control_type="Window")
            move_dialog.wait('visible', timeout=10)

            # Tree container which stores all the tree items
            tree_view = self.outlook.child_window(auto_id="4513", control_type="Tree")

            # Problem with normal condition is:-Every item that is not Tasks that logs an error like not found
            # If there are 10 folders, you will see 9 errors even if the last one is Task it floods your logs.
            # So We initialized task_item as None because we don't found the Tasks folder yet.
            task_item = None
            # .descendants gives all the folder items of tree view
            for item in tree_view.descendants(control_type="TreeItem"):
                # Check if the item name has text tasks then it clicks
                if item.window_text().lower() == "tasks":
                    # Now at this point task item is not None,it now points to the TreeItem object Tasks.
                    task_item = item
                    # Exist the loop after clicking on break
                    break

            # If task item is not none then it makes click otherwise throws error.
            if task_item:
                task_item.click_input()
                self.logger.info("Task clicked")
            else:
                self.logger.error("Task folder not found")
        except Exception as e:
            self.logger.error(f"Failed to click 'Move Items':{e}")
            raise

    #Click on ok/cancel/close
    def click_on_ok_cancel_or_close(self):
        try:
            ok_btn=self.outlook.child_window(title="OK", auto_id="1", control_type="Button")
            ok_btn.click_input()
            self.logger.info("Clicked on OK")
        except Exception as e:
            self.logger.exception(f"Failed to click 'OK' button: {e}")
            try:
                cancel_btn=self.outlook.child_window(title="Cancel", auto_id="2", control_type="Button")
                cancel_btn.click_input()
                self.logger.info("Clicked on Cancel")
            except Exception as e:
                self.logger.exception(f"Failed to click 'Cancel' button: {e}")
                try:
                    close_icon=self.outlook.child_window(title="Close", control_type="Button")
                    close_icon.click_input()
                    self.logger.info("Clicked on Close")
                except Exception as e:
                    self.logger.exception(f"Failed to click 'Close' button:{e}")

    #Set due date
    def set_due_date(self,due_date="12/12/2025"):
        try:
            self.task_window = self.app.window(title_re=".* - Task.*",found_index=0)
            self.task_window.wait('visible', timeout=10)
            self.logger.info("New Task window is opened")

            edit_due_date = self.task_window.child_window(auto_id="4101", control_type="Edit")
            edit_due_date.set_text(due_date)
            self.logger.info(f"Due date set to {due_date}")
        except Exception as e:
            self.logger.error(f"Failed to set due date: {e}")
            #Re-raise/throws the error
            raise

    #Set priority to High
    def set_priority(self):
        try:
            priority_combo = self.task_window.descendants(control_type="ComboBox")[1]
            priority_combo.click_input()
            send_keys("{DOWN}{DOWN}{ENTER}")  #High priority is third is 3rd item
            self.logger.info("Priority set to High")
        except Exception as e:
            self.logger.error(f"Failed to set priority: {e}")
            raise

    #set riminder
    def check_reminder(self):
        try:
            reminder=self.task_window.child_window(title="Reminder", auto_id="4226", control_type="CheckBox")
            reminder.toggle()
            self.logger.info("Reminder checkbox is checked")
        except Exception as e:
            self.logger.error(f"Failed to click 'Reminder' button:{e}")
            raise

    #Set reminder date
    def set_reminder_date(self,reminder_date="30/12/2025"):
        try:
            reminder_due_date=self.task_window.child_window(auto_id="4102", control_type="Edit")
            reminder_due_date.set_text(reminder_date)
            self.logger.info(f"Reminder date set to {reminder_date}")
        except Exception as e:
            self.logger.error(f"Failed to set reminder date:{e}")
            raise

    #Set reminder time
    def set_reminder_time(self):
        try:
            set_time=self.task_window.child_window(title="8:00 AM",auto_id="4108", control_type="Edit")
            set_time.click_input()
            set_time.type_keys("02:30 PM", with_spaces=True)
            self.logger.info("Time set to 02:30 PM")
        except Exception as e:
            self.logger.error(f"Failed to click on time button: {e}")
            raise

    #Click on save and close
    def click_on_save_and_close(self):
        try:
            save_close=self.task_window.child_window(title="Save & Close", control_type="Button")
            save_close.click_input()
            self.logger.info("Clicked on Save and Close")
        except Exception as e:
            self.logger.error(f"Failed to click on save and close button: {e}")
            raise

    #Click on Task menu on main window
    def click_on_task_menu_to_do_list(self):
        try:

            self.click(control_title="Tasks", control_type="TreeItem")
            task_to_do_dlg = self.app.window(title_re=".*Tasks.*")
            task_to_do_dlg.wait('exists visible ready', timeout=10)
            self.logger.info("task_to_do_dlg is opened")
            task_to_do_dlg.print_control_identifiers()

            # Main container for the task for a month
            group_box = task_to_do_dlg.child_window(title="Group By: Expanded:Flag: Due Date: Next Month",control_type="Group")


            # All children of that main container
            task_items = group_box.children(control_type="DataItem")
            self.logger.info(f"{len(task_items)}Task items found")

            if task_items:
                recent_task = task_items[0]
                recent_task.click_input()
                self.logger.info(f"Most recent unread email clicked: {recent_task}")
            else:
                raise Exception("No emails found in the inbox")
        except Exception as e:
            self.logger.error(f"Failed to click 'Tasks' button: {e}")
            raise

    #Click on delete icon.
    def click_on_delete(self):
        self.click_child_window(control_title="Delete", control_type="Button")
























    