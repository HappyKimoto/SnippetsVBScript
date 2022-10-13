@echo off
:: output file name
set "FileName=AlarmReport.txt"
:: get info from the environment and overwrite the output file
echo Computer Name: %ComputerName% > %FileName%
echo Report Created: %DATE% %TIME% >> %FileName%
:: get info from the user
SETLOCAL
CALL :GetUserInput "Machine Name"
CALL :GetUserInput "Alarm Name"
CALL :GetUserInput "Alarm Timestamp"
CALL :GetUserInput "Problem Description"
CALL :GetUserInput "Suspected Cause"
CALL :GetUserInput "Alarm Frequency"
CALL :GetUserInput "Recent Change"
CALL :GetUserInput "Current Machine Condition"
EXIT /B %ERRORLEVEL%
:: function to process user input
:GetUserInput
set /p "UserInput=%~1: "
echo %~1: %UserInput% >> %FileName%
EXIT /B 0