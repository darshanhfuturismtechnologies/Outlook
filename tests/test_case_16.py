from pages.Msg_on_teams_thr_outlook import MsgTeams


def test_msg_on_teams_through_outlook(setup_outlook):
    teams_msg = MsgTeams(setup_outlook)
    teams_msg.click_on_new_item_menu()
    teams_msg.select_chat_option()
    teams_msg.capture_teams_dlg()
    teams_msg.enter_field_to()
    teams_msg.edit_and_send_msg()
    teams_msg.send_msg()
    teams_msg.close_microsoft_teams_app()
