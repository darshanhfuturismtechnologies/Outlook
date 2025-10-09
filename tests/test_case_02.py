
from pages.Unread_mail import UnreadMail

def test_unread_mail(setup_outlook):
    unread_mail=UnreadMail(setup_outlook)
    unread_mail.click_on_unread_mail()
    unread_mail.click_on_recent_unread_mail()