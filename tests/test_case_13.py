from pages.Mark_flag_and_move_in_folder import FlaggedMails


def test_mark_as_flag_and_move_to_folder(setup_outlook):
    flagged_mail=FlaggedMails(setup_outlook)
    flagged_mail.click_on_inbox()
    flagged_mail.search_mail_of_demo_sub()
    flagged_mail.right_click_on_mail()
    flagged_mail.click_on_follow_up()
    flagged_mail.click_on_no_date_flag()
    flagged_mail.click_on_inbox_and_make_right_click()
    flagged_mail.click_on_create_new_folder()
    flagged_mail.enter_folder_name()
    flagged_mail.click_on_demo_folder_via_move()
    flagged_mail.delete_demo_folder()
    flagged_mail.handle_pop_up()
