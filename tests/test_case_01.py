
from pages.ReadMail import ReadMail

#Use the fixture directly in the test function
#We initialize the ReadMail class using the setup_outlook fixture,which is automatically passed to the test function as setup_outlook.

def test_read_mail(setup_outlook):
    # Initialize the ReadMail class with the setup_outlook fixture
    reader = ReadMail(setup_outlook)
    reader.click_on_inbox()
    reader.select_most_recent_email()
    reader.right_click_on_mail()
    reader.mark_as_read_or_unread()


