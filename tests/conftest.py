import pytest
from pywinauto import Application

#Fixtures are used to set up and tear down resources needed for tests.
#Scope=The fixture will be created before each test function and destroyed after the test function finishes.
@pytest.fixture(scope="function")
def setup_outlook():
    #This fixture sets up the Outlook application before each test and tears it down after the test is complete.
    #Application path
    outlook_path = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"

    #Launch application
    app = Application(backend="uia").start(outlook_path)

    #Wait for outlook to be ready by checking the CPU usage drops below a specified threshold 5.0
    app.wait_cpu_usage_lower(threshold=5.0, timeout=30, usage_interval=1.0)

    #Return the app object to the test function for interaction
    yield app

    #Close the Outlook application after each test
    app.kill()


#r means raw string backslashes treated as literal.
#Fixture Setup: Launches Outlook and waits until it is ready.
#Fixture Teardown: Closes the Outlook application to ensure each test is isolated and there’s no leftover state.





