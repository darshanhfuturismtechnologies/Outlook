from pages.Reply import ReplyMail


def test_reply_to_mail(setup_outlook):
    replay_mail=ReplyMail(setup_outlook)
    replay_mail.navigate_towards_inbox()
    replay_mail.click_on_first_mail_of_demo_sub()
    replay_mail.click_on_reply()
    replay_mail.edit_mail_for_reply()
    replay_mail.click_on_send()
