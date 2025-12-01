import pytest

from Helper.helper import Helper
from pages.contact_management import ContactManagement


@pytest.mark.parametrize("test_case",[tc for tc in Helper.load_test_data() if tc["action"] == "Contact_Management"])

def test_contact_management(setup_outlook,test_case):
    contact=ContactManagement(setup_outlook,test_case)
    contact.open_contact_view_from_search_bar()
    contact.add_new_contact()
    contact.enter_contact_details()
    contact.click_on_save_and_close()


