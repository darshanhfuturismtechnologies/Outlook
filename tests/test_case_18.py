from pages.Signature import Signature


def test_signature(setup_outlook):
    sign=Signature(setup_outlook)
    sign.click_on_new_mail()
    sign.click_signature_menu_select_signature_option()
    sign.capture_signatures_and_stationery_window()
    sign.click_on_new()
    sign.capture_new_sign_dlg()
    sign.type_name_for_this_signature()
    sign.new_msg()
    sign.edit_signature()
    sign.click_on_ok()
    sign.get_email_body_text(sign.New_email_dlg)
    sign.close_recent_email_dlg()
    sign.again_click_on_new_mail()