from pages.Search_people import SearchPeople


def test_search_people_and_send_msg_on_teams(setup_outlook):
    search_people=SearchPeople(setup_outlook)
    search_people.click_on_search_people_and_enter_contact()
