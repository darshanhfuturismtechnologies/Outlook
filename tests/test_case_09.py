from pages.Recover_mail import RecoverMail


def test_recover_mail_from_server_backup(setup_outlook):
    recover_mail=RecoverMail(setup_outlook)
    recover_mail.click_on_deleted_items()
    recover_mail.click_on_recover_items_recently_removed()
    recover_mail.click_on_ok()
