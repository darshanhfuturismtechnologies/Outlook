import time
from Helper.helper import Helper

"""TC_06:-
1.Open contact view from search bar
2.Add new contact:Click “New Contact”,Fill in name, phone, email, address and click on Save & close
3.Update contact:Search contact by name,Open and edit details,Save changes
4.Delete contact:Select contact(s),Press Delete or right-click -> Delete"""


class ContactManagement(Helper):
    def __init__(self, app):
        super().__init__(app, ".*Outlook.*")
        self.New_contact_dlg = None

    #Open contact view from search bar
    def open_contact_view_from_search_bar(self):
        tell_me_box = self.outlook.child_window(auto_id="TellMeTextBoxAutomationId", control_type="Edit").wait('visible', timeout=5)
        tell_me_box.type_keys("Contacts")
        self.outlook.child_window(title="Contacts", control_type="MenuItem").click_input()

    def add_new_contact(self):
        contact_window = self.app.window(title_re=".*Contacts.*")
        contact_window.wait('exists visible', timeout=10)

        contact_window.child_window(title="New Contact", control_type="Button").click_input()
        self.logger.info("Clicked on New Contact button")

        New_contact_dlg = self.app.window(title_re=".*- Contact.*")
        New_contact_dlg.wait('visible', timeout=10)
        self.logger.info("New contact dialog opened")
        self.New_contact_dlg = New_contact_dlg  #Save it as an attribute
        #Now you can use the print_control_identifiers method correctly
        # self.New_contact_dlg.print_control_identifiers()

    def enter_contact_details(self):
        Full_name = self.New_contact_dlg.child_window(auto_id="4096", control_type="Edit")
        Full_name.type_keys("Name1{ENTER}")
        Company_name = self.New_contact_dlg.child_window(auto_id="4481", control_type="Edit")
        Company_name.type_keys("Company1{ENTER}")
        Job_title = self.New_contact_dlg.child_window(auto_id="4480", control_type="Edit")
        Job_title.type_keys("Job1{ENTER}")

        Internet_email = self.New_contact_dlg.child_window(auto_id="4120", control_type="Edit")
        Internet_email.type_keys("InternetEmail11@gmail.com{ENTER}")

        Im_address = self.New_contact_dlg.child_window(auto_id="4118", control_type="Edit")
        Im_address.type_keys("ImAddress1{ENTER}")
        self.logger.info("All details are entered")

    def click_on_save_and_close(self):
        self.New_contact_dlg.child_window(title="Save & Close", control_type="Button").click_input()
        self.logger.info("Clicked on Save and Close button")
        time.sleep(3)








