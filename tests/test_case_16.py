import pytest

from Helper.helper import Helper
from pages.Msg_on_teams_thr_outlook import MsgTeams

@pytest.mark.parametrize("test_case",[tc for tc in Helper.load_test_data() if tc["action"] == "Message_on_teams"])

def test_msg_on_teams_through_outlook(setup_outlook,test_case):
    teams_msg = MsgTeams(setup_outlook,test_case)
    teams_msg.click_on_new_item_menu()
    teams_msg.select_chat_option()
    teams_msg.capture_teams_dlg()
    teams_msg.enter_field_to()
    teams_msg.edit_and_send_msg()
    teams_msg.send_msg()
    teams_msg.close_microsoft_teams_app()
