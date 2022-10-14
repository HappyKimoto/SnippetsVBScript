@echo off
:: output file name
set "FileName=AlarmReport.txt"
:: overwrite the output file with Computer Name
echo Computer Name: %ComputerName% > %FileName%
:: append extra system information
echo Report Created: %DATE% %TIME% >> %FileName%
reg query HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation >> %FileName%
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