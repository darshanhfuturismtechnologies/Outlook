import pytest

from Helper.helper import Helper
from pages.Draft_mail import Draft_mail

@pytest.mark.parametrize("test_case",[tc for tc in Helper.load_test_data() if tc["action"] == "Draft_Mail"])

def test_draft_mail(setup_outlook,test_case):
    draft=Draft_mail(setup_outlook,test_case)
    draft.click_on_new_email()
    draft.capture_new_email_dlg()
    draft.enter_recipient()
    draft.enter_cc()
    draft.enter_subject()
    draft.enter_body()
    draft.click_on_close()
    draft.handle_popup_and_click_on_ok()
    draft.click_on_draft()
