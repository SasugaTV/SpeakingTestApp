@echo off
REM Find Duplicate Speaking Tests
REM Finds duplicate student tests and moves older ones to Duplicates folder

echo.
echo ================================================================================
echo Find Duplicate Speaking Tests
echo ================================================================================
echo.

python find_duplicates.py

echo.
echo Press any key to exit...
pause > nul
