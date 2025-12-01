import pytest

from Helper.helper import Helper
from pages.Forward_email import ForwardE


@pytest.mark.parametrize("test_case",[tc for tc in Helper.load_test_data() if tc["action"] == "Forward_Mail"])

def test_forward_email(setup_outlook,test_case):
    forward_email=ForwardE(setup_outlook,test_case)
    forward_email.click_on_inbox()
    forward_email.select_most_recent_mail_of_demo_sub()
    forward_email.click_on_categorized_menu_and_select_category()
    forward_email.click_on_the_recent_mail()
    forward_email.click_on_forward_email()
    forward_email.pop_out_screen()
    forward_email.enter_details()
    forward_email.click_on_send()
    forward_email.clear_search_bar()