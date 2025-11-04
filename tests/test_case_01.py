from pages.Send_mail import SendMail

def test_send_mail(setup_outlook):
    send_mail=SendMail(setup_outlook)
    send_mail.click_on_send_mail()
    send_mail.enter_to_mail_id()
    send_mail.enter_cc_mail_id()
    send_mail.enter_subject()
    send_mail.enter_text_in_body()
    send_mail.attach_file_to_mail()
    send_mail.attach_file_as_copy()
    send_mail.click_on_send()
