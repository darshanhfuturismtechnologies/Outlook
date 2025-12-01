# import time
# from Helper.helper import Helper
#
# class SearchPeople(Helper):
#     def __init__(self,app):
#         super().__init__(app,".*Outlook.*")
#
#     def click_on_search_people_and_enter_contact(self,contact_name):
#         find_group = self.outlook.child_window(title="Find", control_type="Group")
#         search_edit = find_group.child_window(control_type="Edit", found_index=0)
#         search_edit.wait("ready", timeout=10)
#         search_edit.click_input()
#
#         # 2. Type name and press ENTER twice
#         search_edit.type_keys(contact_name + "{ENTER}{ENTER}",with_spaces=True)
#
#         result_list = self.outlook.child_window(control_type="List")
#         results = result_list.children()
#
#         # Check if any result contains the contact name
#         for result in results:
#             if contact_name.lower() in result.window_text().lower():
#                 return result.window_text()  # Return the contact info if found
#
#         # Return None if contact is not found
#         return None

import time
from Helper.helper import Helper

class SearchPeople(Helper):
    def __init__(self, app,test_data):
        super().__init__(app, ".*Outlook.*")
        self.test_data = test_data

    def click_on_search_people_and_enter_contact(self):
        find_group = self.outlook.child_window(title="Find", control_type="Group")
        search_edit = find_group.child_window(control_type="Edit", found_index=0)
        search_edit.wait("ready", timeout=10)
        search_edit.click_input()

        search_edit.type_keys(self.test_data["name"] + "{ENTER}{ENTER}",pause=0.1, with_spaces=True)
        self.logger.info(f"Searched contact: {self.test_data["name"]}")
        time.sleep(2)

        self.logger.info(f"Searched contact: {self.test_data["name"]}")
        time.sleep(2)
        self.logger.info(f"Search contact: {self.test_data["name"]}")
        # self.outlook.print_control_identifiers()


    # if contact:
    #     print(f"Found contact: {contact}")
    # else:
    #     print("Contact not found.")

    # Example usage: Search for a contact named "John Doe"
    # contact = search_person_in_outlook("John Doe")









        # self.outlook. child_window(title="IM", auto_id="PersonaSendIMMenu", control_type="MenuItem").click_input()
        # self.logger.info("Clicked")
        # time.sleep(10)
        #self.outlook.print_control_identifiers()


        # #Locate the Find group box
        # find_group = self.outlook.child_window(title="Find", control_type="Group")
        # #Locate the Search People edit box
        # search_people_edit = find_group.child_window(control_type="Edit",found_index=0)
        # search_people_edit.wait("ready", timeout=5)
        # search_people_edit.click_input()
        # search_people_edit.type_keys("aarti t{ENTER}", with_spaces=True)
        # self.logger.info("Text is entered")

    # def click_on_contact(self):
    #     list_box=self.outlook.child_window(title="Press down arrow for contact suggestions, press enter to select contact.",control_type="List")
    #     list_item=list_box.descendants(control_type="ListItem")
    #
    #     for item in list_item:
    #         if item.window_text() == "Aarti T":
    #             item.click_input()  # click the matching item
    #             self.logger.info(f"Selected contact: {item.window_text()}")
    #             break  # stop looping once found and clicked
    #     else:
    #         self.logger.error("No contacts selected")
    #         self.outlook.print_control_identifiers()
    #
    #     # #
    #     # contact_btn = self.outlook.child_window(title="Aarti T, Available. QA Intern.", control_type="Button")
    #     # contact_btn.click_input()
    #     # self.logger.info("Clicked on contact")
    #     # self.outlook.print_control_identifiers()
    #
    #     time.sleep(20)
