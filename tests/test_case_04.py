
import pytest

from Helper.helper import Helper
from pages.Reply import ReplyMail


#Parametrize send_email tests directly
@pytest.mark.parametrize("test_case",[tc for tc in Helper.load_test_data() if tc["action"] == "Reply_mail"])  # Only send_email tests

def test_reply_mail(setup_outlook, test_case):
    reply_mail = ReplyMail(setup_outlook, test_case)
    reply_mail.navigate_towards_inbox()
    reply_mail.click_on_first_mail_of_demo_sub()
    reply_mail.click_on_reply()
    reply_mail.edit_mail_for_reply()
    reply_mail.click_on_send()
