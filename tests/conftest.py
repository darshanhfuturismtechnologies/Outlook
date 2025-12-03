import datetime
import os
import win32com.client as win32
import pytest
from pywinauto import Application
from Screen_recorder.Screen_recorder import ScreenRecorder

#Empty Dictionary to store recorders
_test_recorders = {}

#Fixture to launch Outlook
#This fixture sets up the Outlook application before each test_recordings and tears it down after the test_recordings is complete.
# Launch Outlook before the test.
# Wait until CPU usage drops → app is fully loaded.
# yield app → gives the running app to the test function.
# After test finishes → app.kill() closes Outlook.

@pytest.fixture(scope="function")
def setup_outlook():
    outlook_path = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
    app = Application(backend="uia").start(outlook_path)
    app.wait_cpu_usage_lower(threshold=5.0, timeout=30, usage_interval=1.0)
    yield app
    app.kill()

# Runs before each test start recording.
# Creates a ScreenRecorder for the current test.
# Starts recording immediately.
# Saves the recorder in _test_recorders using item.nodeid.
# yield:-lets the actual test execute while recording continues.

@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_setup(item):
    project_folder = r"C:\Users\swapnalik\PycharmProjects\Outlook\test_recordings"
    system_folder = r"C:\Fail_test_recordings\Outlook_failure"

    recording = ScreenRecorder(f"{item.name}.mp4", project_folder, system_folder)
    recording.start()
    _test_recorders[item.nodeid] = recording
    yield  # let test run

# Capture test result (pass/fail).
# It Runs after the test function executes.
#result.when == "call" ensures we only check the actual test execution not setup/teardown
#Stores whether the test failed or passed in item._test_failed.

@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_makereport(item):
    outcome = yield
    result = outcome.get_result()
    if result.when == "call":
        item._test_failed = result.failed


# Runs after the test and all teardown steps.
# Retrieves the recorder for this test from _test_recorders.
# Checks if the test failed (_test_failed).
# Calls recorder.stop.
# If test failed:-video is saved.
# If test passed :-video is discarded.

@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_teardown(item):
    """Stop recording and save only if test failed."""
    yield
    recorder = _test_recorders.pop(item.nodeid,None)
    if recorder:
        save_video = getattr(item, "_test_failed", False)
        recorder.stop(save=save_video)




# """Auto trigger email configuration for sending reports after test execuion"""



# REPORT_DIR = r"C:\Users\swapnalik\PycharmProjects\Outlook\Html_reports"
# Email_receiver = "swapnalik@futurismtechnologies.com"
# Email_subject = "Automated Test Execution Report"
# Email_body=(
#     "Hello,<br><br>"
#     "Please find attached the detailed HTML report for the executed test cases.<br><br>"
#     "Best regards,<br>"
#     "Swapnali K,<br><br>"
# )
#
# # Ensure report directory exists
# os.makedirs(REPORT_DIR, exist_ok=True)
#
# def send_email_outlook(report_file,receiver):
#     """
#     Sends an email with the given report file attached via Outlook.
#     """
#     try:
#         outlook = win32.Dispatch('Outlook.Application')
#         mail = outlook.CreateItem(0)
#         mail.To = receiver
#         mail.Subject = Email_subject
#         mail.HTMLBody = Email_body
#
#         if os.path.exists(report_file):
#             mail.Attachments.Add(Source=report_file)
#             print(f"Attached report:{report_file}")
#         else:
#             print("Report file not found,cannot attach.")
#
#         mail.Send()
#         print(f"Email sent successfully to {receiver}.")
#
#     except Exception as e:
#         print(f"Failed to send email: {e}")
#
#
# def pytest_configure(config):
#     """
#     Set HTML report path and format before tests run.
#     """
#     global REPORT_FILE
#     REPORT_FILE = os.path.join(REPORT_DIR,f"report_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.html")
#     config.option.htmlpath = REPORT_FILE
#     config.option.self_contained_html = True
#     print(f"HTML report will be saved at: {REPORT_FILE}")
#
#
# def pytest_unconfigure(config):
#     """
#     Called after all tests and plugins are finished.
#     Ensures HTML report is complete before emailing.
#     """
#     if os.path.exists(REPORT_FILE):
#         if os.path.getsize(REPORT_FILE) > 100:
#             print(f"Final HTML report ready: {REPORT_FILE}")
#             send_email_outlook(report_file=REPORT_FILE, receiver=Email_receiver)
#         else:
#             print("Report file is empty. Email not sent.")
#     else:
#         print("Report file not generated. Email not sent.")

























# import pytest
# from pywinauto import Application
# from Screen_recorder.Screen_recorder import ScreenRecorder
#
#
# #This fixture sets up the Outlook application before each test_recordings and tears it down after the test_recordings is complete.
# @pytest.fixture(scope="function")
# def setup_outlook():
#     outlook_path =r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
#     #Launch application
#     app = Application(backend="uia").start(outlook_path)
#     #Wait for outlook to be ready by checking the CPU usage drops below a specified threshold 5.0
#     app.wait_cpu_usage_lower(threshold=5.0, timeout=30, usage_interval=1.0)
#     #Return the app object to the test_recordings function for interaction
#     yield app
#     #Close the Outlook application after each test_recordings
#     app.kill()
#
#
# """Hook — Start Recording When a Test Fails"""
# #hookwrapper=True:-Allows us to run code before and after Pytest collects the test_recordings report.
# #outcome = yield:-Wait for the original Pytest test_recordings report to be generated.
# #result = outcome.get_result():-Get the report object, which contains test_recordings results (passed,failed,skipped).
# # result.when == "call" → Check only the main test_recordings execution, not setup/teardown.
# # result.failed → Check if the test_recordings actually failed.
# # ScreenRecorder(f"{item.name}_failure.mp4") → Create a recorder for this failed test_recordings.
# # recorder.start() → Start recording immediately.
# # _failure_recorders[item.nodeid] = recorder → Store the recorder in the dictionary.
# # print(...) → Debug/log info to console.
# #item.nodeid (unique test_recordings ID),Value = ScreenRecorder instance.
# #Needed because multiple tests might fail, and you want to record each separately
#
#
# #Dictionary to store recorders for each failing test_recordings.
# _failure_recorders = {}
#
# @pytest.hookimpl(hookwrapper=True)
# def pytest_runtest_makereport(item):
#     outcome = yield
#     result = outcome.get_result()
#
#     if result.when == "call" and result.failed:
#         project_folder = r"C:\Users\swapnalik\PycharmProjects\Outlook\test_recordings"
#         system_folder = r"C:\Fail_test_recordings\Outlook_failure"
#
#         #Start recording for the failed test case (in both locations)
#         recorder = ScreenRecorder(f"{item.name}_failure.mp4",project_folder, system_folder)
#         recorder.start()
#         _failure_recorders[item.nodeid] = recorder
#         print(f"Started failure recording for {item.nodeid}")
#
#
# """Hook — Stop Failure Recording After Teardown"""""
# #Runs after each test_recordings’s teardown phase — even if it failed.
# #Retrieves and removes (pop) the corresponding recorder from the dictionary.
# #If a recorder exists,it stops the recording.
# #_failure_recorders.pop(item.nodeid, None):-Get the recorder for this test_recordings and remove it from the dictionary.
# #recorder.stop():-Stop the recording after teardown finishes.
#
# @pytest.hookimpl(hookwrapper=True)
# def pytest_runtest_teardown(item):
#     yield
#     recorder = _failure_recorders.pop(item.nodeid,None)
#     if recorder:
#         recorder.stop()
#         print(f"Stopped failure recording for {item.nodeid}")






















# import pytest
# from pywinauto import Application
#
# #Fixtures are used to set up and tear down resources needed for tests.
# #Scope=The fixture will be created before each test_recordings function and destroyed after the test_recordings function finishes.
# @pytest.fixture(scope="function")
# def setup_outlook():
#     #This fixture sets up the Outlook application before each test_recordings and tears it down after the test_recordings is complete.
#     #Application path
#     outlook_path = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
#
#     #Launch application
#     app = Application(backend="uia").start(outlook_path)
#
#     #Wait for outlook to be ready by checking the CPU usage drops below a specified threshold 5.0
#     app.wait_cpu_usage_lower(threshold=5.0, timeout=30, usage_interval=1.0)
#
#     #Return the app object to the test_recordings function for interaction
#     yield app
#
#     #Close the Outlook application after each test_recordings
#     app.kill()
#
#
# #r means raw string backslashes treated as literal.
# #Fixture Setup: Launches Outlook and waits until it is ready.
# #Fixture Teardown: Closes the Outlook application to ensure each test_recordings is isolated and there’s no leftover state.