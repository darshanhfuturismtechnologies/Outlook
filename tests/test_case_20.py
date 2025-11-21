from pages.Reading_pane import ReadingPane


def test_reading_pane(setup_outlook):
    reading_pane=ReadingPane(setup_outlook)
    reading_pane.navigate_to_view_page()
    reading_pane.click_on_reading_pane()
    reading_pane.switch_to_bottom_reading_pane()