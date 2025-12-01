import pytest

from Helper.helper import Helper
from pages.Download_attatchments import Download_attachments

@pytest.mark.parametrize("test_case",[tc for tc in Helper.load_test_data() if tc["action"] == "Download_Attachments"])

def test_download_attachment_from_email(setup_outlook,test_case):
    attachment = Download_attachments(setup_outlook,test_case)
    attachment.search_mail_of_demo_sub()
    attachment.click_on_attachment_option()
    attachment.click_on_save_attachment()
    attachment.save_attachment_at_folder()
    attachment.save_file_with_unique_name()
    attachment.edit_file_name()
    attachment.click_on_save()

