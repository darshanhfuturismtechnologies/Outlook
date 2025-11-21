from pages.Quick_step import QuickSteps


def test_quick_steps(setup_outlook):
    quick_steps = QuickSteps(setup_outlook)
    quick_steps.click_on_create_quick_steps()
    quick_steps.capture_quick_steps_dlg()
    quick_steps.choose_an_action()
    quick_steps.choose_short_cut_key()
    quick_steps.click_on_finish_by_handling_pop_up()
    quick_steps.select_mail_mark_as_complete_on_demo_sub_mail()
    quick_steps.click_on_mark_ss_complete_inside_quick_steps()
    quick_steps.delete_mark_complete_quick_step()
    quick_steps.click_on_delete()