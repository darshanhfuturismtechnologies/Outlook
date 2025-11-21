import time
from Helper.helper import Helper


class ChangeView(Helper):
    def __init__(self,app):
        super().__init__(app,".*Outlook.*")
        self.change_view = None
        self.view_menu = None

    def click_on_view(self):
        self.click(control_title="View",control_type="TabItem")

    def click_on_change_view(self):
        self.change_view=self.outlook.child_window(title="Change View", auto_id="CurrentViewGallery", control_type="MenuItem")
        self.change_view.click_input()
        self.logger.info("Clicked on Change View")

        self.view_menu = self.change_view.child_window(title="Change View",control_type="Menu",found_index=0)
        self.view_menu.wait('visible exists',timeout=5)
        self.logger.info("View menu is visible")
        # self.outlook.print_control_identifiers()

    def change_into_compact_view(self):
        try:
            compact_view=self.view_menu.child_window(title="Compact", control_type="ListItem")
            compact_view.click_input()
            self.logger.info("Clicked on compact view")
            time.sleep(5)
        except Exception as e:
            self.logger.error(f"Failed to click on compact view:{e}")

    def change_into_single_view(self):
        try:
            self.click_on_change_view()
            single_view = self.view_menu.child_window(title="Single", control_type="ListItem")
            single_view.click_input()
            self.logger.info("Clicked on single view")
            time.sleep(5)
        except Exception as e:
            self.logger.error(f"Failed to click on single view:{e}")


    def change_into_preview_view(self):
        try:
            self.click_on_change_view()
            preview_view = self.view_menu.child_window(title="Preview", control_type="ListItem")
            preview_view.click_input()
            self.logger.info("Clicked on preview view")
            time.sleep(5)
        except Exception as e:
            self.logger.error(f"Failed to click on preview view:{e}")

    def change_into_default_view(self):
        try:
            self.click_on_change_view()
            default_view = self.view_menu.child_window(title="Sent To", control_type="ListItem")
            default_view.click_input()
            self.logger.info("Clicked on default view")
            time.sleep(5)
        except Exception as e:
            self.logger.error(f"Failed to click on default view:{e}")


