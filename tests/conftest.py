
import pytest
from pywinauto import Application
from Screen_recorder.Screen_recorder import ScreenRecorder


#scope=session: the fixture runs once per test session, not per test.
#autouse=True: automatically used by all tests without needing to explicitly add it to test functions.
@pytest.fixture(scope="session", autouse=True)
def session_record():
    """Start recording for the whole test session."""
    recorder = ScreenRecorder("tests_record.mp4")
    recorder.start()
    yield
    recorder.stop()


#This fixture sets up the Outlook application before each test and tears it down after the test is complete.
@pytest.fixture(scope="function")
def setup_outlook():
    outlook_path = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"

    #Launch application
    app = Application(backend="uia").start(outlook_path)

    #Wait for outlook to be ready by checking the CPU usage drops below a specified threshold 5.0
    app.wait_cpu_usage_lower(threshold=5.0, timeout=30, usage_interval=1.0)

    #Return the app object to the test function for interaction
    yield app

    #Close the Outlook application after each test
    app.kill()


#Dictionary to store recorders for each failing test.
_failure_recorders = {}

"""Hook called after each test phase — start recording on failure."""
#hookwrapper=True:-Allows us to run code before and after Pytest collects the test report.
#outcome = yield:-Wait for the original Pytest test report to be generated.
#result = outcome.get_result():-Get the report object, which contains test results (passed,failed,skipped).
# result.when == "call" → Check only the main test execution, not setup/teardown.
# result.failed → Check if the test actually failed.
# ScreenRecorder(f"{item.name}_failure.mp4") → Create a recorder for this failed test.
# recorder.start() → Start recording immediately.
# _failure_recorders[item.nodeid] = recorder → Store the recorder in the dictionary.
# print(...) → Debug/log info to console.
#item.nodeid (unique test ID),Value = ScreenRecorder instance.
#Needed because multiple tests might fail, and you want to record each separately

@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_makereport(item):
    outcome = yield
    result = outcome.get_result()

    if result.when == "call" and result.failed:
        recorder = ScreenRecorder(f"{item.name}_failure.mp4")
        recorder.start()
        _failure_recorders[item.nodeid] = recorder
        print(f"Started failure recording for {item.nodeid}")



"""Runs after each test, even if the test failed."""
#_failure_recorders.pop(item.nodeid, None):-Get the recorder for this test and remove it from the dictionary.
#recorder.stop():-Stop the recording after teardown finishes.

@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_teardown(item):
    """Stop failure recording after test teardown."""
    yield
    recorder = _failure_recorders.pop(item.nodeid, None)
    if recorder:
        recorder.stop()
        print(f"Stopped failure recording for {item.nodeid}")



















# import pytest
# from pywinauto import Application
#
# #Fixtures are used to set up and tear down resources needed for tests.
# #Scope=The fixture will be created before each test function and destroyed after the test function finishes.
# @pytest.fixture(scope="function")
# def setup_outlook():
#     #This fixture sets up the Outlook application before each test and tears it down after the test is complete.
#     #Application path
#     outlook_path = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
#
#     #Launch application
#     app = Application(backend="uia").start(outlook_path)
#
#     #Wait for outlook to be ready by checking the CPU usage drops below a specified threshold 5.0
#     app.wait_cpu_usage_lower(threshold=5.0, timeout=30, usage_interval=1.0)
#
#     #Return the app object to the test function for interaction
#     yield app
#
#     #Close the Outlook application after each test
#     app.kill()
#
#
# #r means raw string backslashes treated as literal.
# #Fixture Setup: Launches Outlook and waits until it is ready.
# #Fixture Teardown: Closes the Outlook application to ensure each test is isolated and there’s no leftover state.