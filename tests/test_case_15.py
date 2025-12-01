import pytest

from Helper.helper import Helper
from pages.Create_rule import Create_rule

@pytest.mark.parametrize("test_case",[tc for tc in Helper.load_test_data() if tc["action"] == "Create_Rule"])

def test_create_rule_and_delete_it(setup_outlook,test_case):
    rule=Create_rule(setup_outlook,test_case)
    rule.navigate_to_file_menu()
    rule.click_on_rule_and_alert()
    rule.rule_and_alert_window()
    rule.click_on_create_rule()
    rule.handle_rules_wizard_window()
    rule.set_conditions_inside_list()
    rule.edit_rule_one()
    rule.edit_rule_two()
    rule.edit_rule_three()
    rule.uncheck_process_rule_and_click_next()
    rule.click_on_next_btn_four()
    rule.click_on_finish_btn()
    rule.handle_microsoft_popup()
    rule.apply_rule_and_click_ok()