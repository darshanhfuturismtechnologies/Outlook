from Helper.helper import Helper


class Forward_email(Helper):
    def __init__(self, app):
        super().__init__(app, ".*Outlook.*")