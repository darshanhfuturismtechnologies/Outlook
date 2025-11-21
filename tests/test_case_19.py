from pages.Change_view import ChangeView


def test_change_view(setup_outlook):
    view=ChangeView(setup_outlook)
    view.click_on_view()
    view.click_on_change_view()
    view.change_into_compact_view()
    view.change_into_single_view()
    view.change_into_preview_view()
    view.change_into_default_view()