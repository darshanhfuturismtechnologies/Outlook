from pages.Delete_mail import DeleteMail


def test_delete_mail(setup_outlook):
    dlt_mail=DeleteMail(setup_outlook)
    dlt_mail.click_on_sent_item()
    dlt_mail.click_on_demo_sub_mail()
    dlt_mail.right_click_on_mail()
    dlt_mail.click_on_delete_mail()
