import time
from pywinauto import Application
from Logger import logs

class Helper:
    def __init__(self,app: Application,title_re:str):                            #Helper constructor called when object is created.
        self.selected_mail = None
        self.app = app                                                          #Store the passed application object so methods can use
        self.logger = logs.logging                                              #Access the logger to record actions.
        self.outlook = self.app.window(title_re=title_re)                       #store the window title
        self.outlook.maximize()
        self.outlook.wait('exists ready visible enabled',timeout=30)            #Waits until the window exists,ready for interaction,visible,and enabled.

    """Method for click element"""
    def click(self,control_title:str,control_type:str):
        #It looks inside window (self.window) for controls that matches the title and control type.
        control = self.outlook.child_window(title=control_title, control_type=control_type)
        try:
            #First checks the control exists/ready or not.
            if control.exists():
                #It clicks input
                control.click_input()                                        # use click_input for better compatibility it Simulates a real physical mouse click (like a human clicking).
            else:
                #Otherwise raise control not found exception
                raise Exception(f'{control_title} not found')
        except Exception as e:
            #If any error happens while checking or clicking,it raises an error with details assign this error to var e.
            raise Exception(f'control {control_title} not found: {e}')
        self.logger.info(f'{control_title} clicked')



    """Click a child control under a given parent window."""
    def click_inside_parent_win(self,parent,title=None,title_re=None,control_type=None,auto_id=None,timeout=10):

        control=parent.child_window(title=title,title_re=title_re,control_type=control_type,auto_id=auto_id)
        control.wait("visible",timeout=timeout)

        try:
            control.click_input()
            self.logger.info(
                f"Clicked child control: title={title or title_re}")
            return control

        except Exception as e:
            self.logger.error(
                f"Failed to click child control: title={title or title_re}")
            raise e


    """This method tries to click on child window inside current active window."""
    def click_child_window(self, control_title: str, control_type: str):
        try:
            #Gets the currently active window
            active_window = self.app.active()
            #Store the child control within the active window
            control = active_window.child_window(title=control_title, control_type=control_type)
            control.click_input()
            self.logger.info(f"Clicked on {control_title}")
        except Exception as e:
            self.logger.error(f"Failed to click on {control_title}: {e}")



    """Click on list item and attach file to the mail"""
    #Parent window contains children as list item
    def click_list_item_by_text(self,parent_container,target_name):
        #Get all child elements of type ListItem under the given parent window
        list_items = parent_container.children(control_type="ListItem")

        #Loop through each ListItem for item
        for item in list_items:
            #Check if the current item's visible text matches the target name then click
            if item.window_text() == target_name:
                item.click_input()
                self.logger.info(f"Clicked on item: {target_name}")
                return True                                                 #True if the item was found and clicked
        # If no matching item was found after the loop then print error
        self.logger.error(f"Item '{target_name}' not found.")
        return False                                                        #Return false to indicate item is not found



    """Wait until email is to be sent and then close the specific window"""
    # Email takes 30-40 seconds to sent and then it closes window
    def wait_for_window_to_close(self,window,timeout=60):
        try:
            #Wait until the window disappears,If window closes within time it returns log success msg
            window.wait_not('visible', timeout=timeout)
            #time.sleep(30)                                                    #Extra time to close window
            self.logger.info(f"Window '{window}' closed successfully.")
        except Exception as e:
            #Otherwise it throws exception like that tells email is not sent
            self.logger.warning(f"Timeout waiting for window to close. Mail might not be sent yet.Error:{e}")



    """Clicks a menu item in Outlook by its title."""
    def click_menu_item(self,menu_title: str,control_type:str ="TreeItem",timeout=5):
        try:
            #Search for the menu item based on the title and control type
            menu_item = self.outlook.child_window(title_re=f".*{menu_title}.*",control_type=control_type,found_index=0)
            #Wait for the menu item to become visible
            menu_item.wait('visible',timeout=timeout)
            # Click the menu item
            menu_item.click_input()
            #Log the successful click
            self.logger.info(f"Successfully clicked on '{menu_title}' menu item.")
            return True
        except Exception as e:
            # Log the error if the item couldn't be clicked
            self.logger.error(f"Failed to click on '{menu_title}' menu item: {e}")
            return False


    """Right-click on the selected email."""
    def right_click_on_selected_option(self):
        try:
            if self.selected_mail:
                self.selected_mail.right_click_input()
                time.sleep(3)
                self.logger.info("Right-clicked on the selected option.")
                return True
            else:
                self.logger.error("No option selected after right-click.")
                return False
        except Exception as e:
            self.logger.error(f"Failed to right-click on option: {e}")
            return False


    """Waits for the context menu to appear and returns it."""
    def get_context_menu(self,timeout=10):
        try:
            context_menu = self.outlook.child_window(title="Context Menu", control_type="Menu")
            context_menu.wait('visible', timeout=timeout)
            return context_menu
        except Exception as e:
            self.logger.error(f"Context menu not found: {e}")
            raise


    """window-wait helper for top-level Outlook windows."""
    def wait_for_window(self,title_re=None, control_type="Window",timeout=10):
        try:
            #Create window object
            window = self.app.window(title_re=title_re,control_type=control_type)
            window.wait("visible", timeout=timeout)
            self.logger.info(f"Window appeared: {title_re}")
            return window

        except Exception as e:
            self.logger.error(f"Window NOT found: {title_re}")
            raise e


    """Wait for a child window under a given parent."""
    def wait_for_child_window(self, parent,title_re=None,control_type=None,timeout=10):

        child = parent.child_window(title_re=title_re,control_type=control_type)

        try:
            child.wait("visible", timeout=timeout)
            self.logger.info(f"Child window visible: title={title_re}")
            return child

        except Exception as e:
            self.logger.error(
                f"Child window NOT found under parent.\n"
                f"title_re={title_re},control_type={control_type}"
            )
            raise e



    """Wait for a control to appear after the view change."""
    def wait_for_view_change(self, control_title, timeout=5):
        try:
            element = self.outlook.child_window(title=control_title, control_type="Button")
            element.wait('visible', timeout=timeout)  # Wait for the element to become visible
            return True
        except Exception as e:
            self.logger.error(f"Failed to detect view change: {e}")
            return False





