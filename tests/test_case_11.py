import pytest

from Helper.helper import Helper
from pages.Group_creation import GroupMail

@pytest.mark.parametrize("test_case",[tc for tc in Helper.load_test_data() if tc["action"] == "Create_Group"])
def test_create_group_and_send_send_mail(setup_outlook,test_case):
    grp_mail=GroupMail(setup_outlook,test_case)
    grp_mail.click_on_search_bar_and_enter_contact()
    grp_mail.click_on_new_group()
    grp_mail.enter_grp_name()
    grp_mail.click_on_add_members()
    grp_mail.select_members()
    grp_mail.click_on_ok_or_cancel()
    grp_mail.click_on_save_and_close()
    grp_mail.click_on_new_item()
    grp_mail.enter_group_name_in_to_field()
    grp_mail.enter_subject()
    grp_mail.enter_text_in_body()
    grp_mail.click_on_send()
