import pytest
from Helper.helper import Helper
from pages.Send_mail import SendMail

#Parametrize send_email tests directly
#It tells pytest to run the test function multiple times, once for each item in the list you provide.
@pytest.mark.parametrize("test_case",[tc for tc in Helper.load_test_data() if tc["action"] == "send_email"])  # Only send_email tests

def test_send_mail(setup_outlook, test_case):
    send_mail = SendMail(setup_outlook, test_case)
    send_mail.click_on_send_mail()
    send_mail.enter_to_mail_id()
    send_mail.enter_cc_mail_id()
    send_mail.enter_subject()
    send_mail.enter_text_in_body()
    send_mail.attach_file_to_mail()
    send_mail.attach_file_as_copy()
    send_mail.click_on_send()
