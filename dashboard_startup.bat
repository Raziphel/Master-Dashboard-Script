@echo off

REM DASH-STARTUP - Batch script to pull the latest PowerShell dashboard script from a network share and run it

REM Define variables
set "SCRIPT_PATH=\\int3\shared\BlakeC\master_dashboard_script.ps1"  REM Path to the shared network location where the PowerShell script is stored
set "LOCAL_PATH=C:\scripts\master_dashboard_script.ps1"  REM Local path to store the downloaded PowerShell script

REM Create the local scripts directory if it doesn't exist
if not exist "C:\scripts" mkdir "C:\scripts"

REM Attempt to copy the script from the network location to the local machine
copy "%SCRIPT_PATH%" "%LOCAL_PATH%" /Y >nul 2>nul

REM Run the script with PowerShell, even if the copy failed
powershell -ExecutionPolicy Bypass -File "%LOCAL_PATH%"

REM Exit the batch file
exit
