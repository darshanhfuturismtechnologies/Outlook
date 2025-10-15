from pages.Move_email_to_folder import MoveToArchiveFolder


def test_move_to_archive(setup_outlook):
    archive=MoveToArchiveFolder(setup_outlook)
    archive.click_on_inbox()
    archive.search_mail()
    archive.right_click_on_mail()
    archive.click_on_move()
    archive.click_on_archive()
