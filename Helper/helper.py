import time

from pywinauto import Application
from Logger import logs

class Helper:
    def __init__(self,app:Application,title_re:str):
        self.send_mail_window = None
        self.app = app                                                          #Store the passed application object
        self.logger = logs.logging
        self.outlook = self.app.window(title_re=title_re)                       #store the window title
        self.outlook.maximize()
        self.outlook.wait('exists',timeout=30)                                  #Wait until window become visible

    """Method for click element"""
    def click(self, control_title: str, control_type: str):
        #It looks inside window (self.window) for controls that matches the title and control type.
        control = self.outlook.child_window(title=control_title, control_type=control_type)
        try:
            #First checks the control exists/ready or not.
            if control.exists():
                #It click input
                control.click_input()                                        # use click_input for better compatibility it Simulates a real physical mouse click (like a human clicking).
            else:
                #Otherwise raise control not found exception
                raise Exception(f'control {control_title} not found')
        except Exception as e:
            #If any error happens while checking or clicking,it raises an error with details assign this error to var e.
            raise Exception(f'control {control_title} not found: {e}')
        self.logger.info(f'control {control_title} clicked')


    """This method tries to click on child window inside current active window."""
    def click_child_window(self, control_title: str, control_type: str):
        try:
            #Retrieve the currently active window
            active_window = self.app.active()
            #Store the child control within the active window
            control = active_window.child_window(title=control_title, control_type=control_type)
            control.click_input()
            self.logger.info(f"Clicked on {control_title}")
        except Exception as e:
            self.logger.error(f"Failed to click on {control_title}: {e}")

    """Click on list item and attach file to the mail"""
    #Parent window contains childrens as listitem
    def click_list_item_by_text(self, parent_container,target_name):
        #Get all child elements of type ListItem under the given parent window
        list_items = parent_container.children(control_type="ListItem")

        #Loop through each ListItem for item
        for item in list_items:
            #Check if the current item's visible text matches the target name then click
            if item.window_text() == target_name:
                item.click_input()
                self.logger.info(f"Clicked on item: {target_name}")
                return True                                             #True if the item was found and clicked
        # If no matching item was found after the loop then print error
        self.logger.error(f"Item '{target_name}' not found.")
        return False                                                     #Return false to indicate item is not found

    """Wait until email is to be sent and then close the specific window"""
    # Email takes 30-40 seconds to sent and then it close window
    def wait_for_window_to_close(self, window,timeout=60):
        try:
            #Wait until the window disappears ,If window closes within time it returns log success msg
            window.wait_not('visible', timeout=timeout)
            time.sleep(30)
            self.logger.info(f"Window '{window}' closed successfully.")
        except Exception as e:
            #Otherwise it throws exception like that tells email is not sent
            self.logger.warning(f"Timeout waiting for window to close. Mail might not be sent yet. Error: {e}")














