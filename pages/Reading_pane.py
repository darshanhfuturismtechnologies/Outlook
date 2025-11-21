import time
from Helper.helper import Helper


class ReadingPane(Helper):
    def __init__(self,app):
        super().__init__(app,".*Outlook.*")
        self.menu_items = None
        self.reading_pane_menu = None
        self.reading_pane_icon = None

    def navigate_to_view_page(self):
        self.click(control_title="View",control_type="TabItem")

    def click_on_reading_pane(self):
        self.reading_pane_icon=self.outlook.child_window(title="Reading Pane",control_type="MenuItem")
        self.reading_pane_icon.click_input()
        self.logger.info("Reading Pane clicked")

        self.reading_pane_menu = self.reading_pane_icon.child_window(title="Reading Pane", control_type="Menu")
        self.reading_pane_menu.wait('visible exists',timeout=5)
        self.logger.info("Reading pane menu is visible")

        self.menu_items = self.reading_pane_menu.descendants(control_type="MenuItem")
        self.logger.info(f"Found {len(self.menu_items)} menu items")

    def switch_to_bottom_reading_pane(self):

        #Switch to Bottom
        try:
            clicked = False
            for item in self.menu_items:
                if item.window_text() == "Bottom":
                    item.click_input()
                    time.sleep(5)
                    self.logger.info("Clicked on Bottom")
                    clicked = True
                    break

            if not clicked:
                self.logger.error("Bottom option not found in Reading Pane menu")

        except Exception as e:
            self.logger.error(f"Error while clicking Bottom: {e}")

        #Switch to Off
        try:
            self.click_on_reading_pane()
            clicked = False
            for item in self.menu_items:
                if item.window_text() == "Off":
                    item.click_input()
                    time.sleep(5)
                    self.logger.info("Clicked on Off")
                    clicked = True
                    break

            if not clicked:
                self.logger.error("Off option not found in Reading Pane menu")

        except Exception as e:
            self.logger.error(f"Error while clicking Off:{e}")

        #Switch to right
        try:
            self.click_on_reading_pane()
            clicked = False
            for item in self.menu_items:
                if item.window_text() == "Right":
                    item.click_input()
                    time.sleep(5)
                    self.logger.info("Clicked on Right")
                    clicked = True
                    break

            if not clicked:
                self.logger.error("Right option not found in Reading Pane menu")

        except Exception as e:
            self.logger.error(f"Error while clicking Right: {e}")





