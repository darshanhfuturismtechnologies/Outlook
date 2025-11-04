

from pages.Convert_mail_to_task import TaskConverter

def test_convert_mail_into_task(setup_outlook):
    task=TaskConverter(setup_outlook)
    task.click_on_deleted_items()
    task.click_on_demo_sub_mail()
    task.right_click_on_mail()
    task.click_on_move()
    task.move_item_to_task()
    task.click_on_ok_cancel_or_close()
    task.set_due_date()
    task.set_priority()
    task.check_reminder()
    task.set_reminder_date()
    task.set_reminder_time()
    task.click_on_save_and_close()
    task.click_on_task_menu_to_do_list()
    task.click_on_delete()