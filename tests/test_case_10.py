import pytest

from Helper.helper import Helper
from pages.Search_people import SearchPeople

@pytest.mark.parametrize("test_case",[tc for tc in Helper.load_test_data() if tc["action"] == "Search_Contact"])

def test_search_people_and_send_msg_on_teams(setup_outlook,test_case):
    search_people=SearchPeople(setup_outlook,test_case)
    search_people.click_on_search_people_and_enter_contact()
