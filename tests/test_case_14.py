from pages.Download_attatchments import Download_attachments


def test_download_attachment_from_email(setup_outlook):
    attachment = Download_attachments(setup_outlook)
    attachment.search_mail_of_demo_sub()
    attachment.click_on_attachment_option()
    attachment.click_on_save_attachment()
    attachment.save_attachment_at_folder()
    attachment.save_file_with_unique_name()
    attachment.edit_file_name()
    attachment.click_on_save()

