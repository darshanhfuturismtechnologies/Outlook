import pytest
from pywinauto import Application

#The fixture will be created before each test function and destroyed after the test function finishes.
@pytest.fixture(scope="function")
def setup_outlook():
    #This fixture sets up the Outlook application before each test and tears it down after the test is complete.
    #Application path
    outlook_path = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"

    #Launch application
    app = Application(backend="uia").start(outlook_path)

    #Wait for outlook to be ready
    app.wait_cpu_usage_lower(threshold=5.0, timeout=30, usage_interval=1.0)

    #Return the app object to the test function for interaction
    yield app

    #Close the Outlook application after each test
    app.kill()





