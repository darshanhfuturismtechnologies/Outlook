@echo off
echo === Running Outlook Test Cases ===

:: Go to your project folder
cd /d "C:\Users\swapnalik\PycharmProjects\Outlook"

:: Set date variable and clean invalid filename characters
set "CURRDATE=%DATE:/=-%"

:: Run pytest on the tests folder, output logs to your Logger folder
python -m pytest tests > "C:\Users\swapnalik\PycharmProjects\Outlook\Logger\scheduled_test_log_%CURRDATE%.txt" 2>&1

echo === Test Execution Completed ===
pause
