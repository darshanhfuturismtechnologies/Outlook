from Helper.helper import Helper

class RecoverMail(Helper):
    def __init__(self,app):
        super().__init__(app,".*Outlook.*")

    def click_on_deleted_items(self):
        self.click_menu_item(".*Deleted Items:.*")
        self.logger.info("Clicked on Deleted Items")

    def click_on_recover_items_recently_removed(self):
        self.click(control_title="Recover items recently removed from this folder", control_type="Button")

        recover_dlt_window=self.outlook.window(title_re=".*Recover Deleted Items.*",control_type="Window")
        recover_dlt_window.wait('visible',timeout=10)
        self.logger.info("Recover deleted items window is visible")

        list_box=recover_dlt_window.child_window(control_type="List")
        list_item=list_box.descendants(control_type="ListItem")
        assert list_item,"No deleted items found to recover"

        list_item[0].click_input()
        self.logger.info("Clicked on first item")

    def click_on_ok(self):
        self.click(control_title="Ok",control_type="Button")
        # self.outlook.print_control_identifiers()



