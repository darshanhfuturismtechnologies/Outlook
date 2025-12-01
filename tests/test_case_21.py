from pages.Change_Fonts import ChangeFont


def test_change_font(setup_outlook):
    font=ChangeFont(setup_outlook)
    font.click_on_file()
    font.click_on_mail()
    font.click_on_stationery_and_font()
    font.capture_signature_and_stationery_dlg()
    font.select_first_font()
    font.edit_font()
