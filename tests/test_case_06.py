from pages.contact_management import ContactManagement

def test_contact_management(setup_outlook):
    contact=ContactManagement(setup_outlook)
    contact.open_contact_view_from_search_bar()
    contact.add_new_contact()
    contact.enter_contact_details()
    contact.click_on_save_and_close()


