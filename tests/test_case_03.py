from pages.ReadMail import ReadMail


def test_read_mail(setup_outlook):
    read=ReadMail(setup_outlook)
    read.select_most_recent_email()
    read.right_click_on_mail()
    read.mark_as_read_or_unread()