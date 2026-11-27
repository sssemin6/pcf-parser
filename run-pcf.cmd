@echo off
setlocal EnableExtensions DisableDelayedExpansion

if exist "%~dp0pcf-runtime\python.exe" goto bundled
where py.exe >nul 2>&1
if not errorlevel 1 goto python_launcher
where python.exe >nul 2>&1
if not errorlevel 1 goto python_command

echo Python 3.10 or later is required. Install Python or use the portable PCF Parser ZIP.
set "pcf_exit=1"
goto finish

:bundled
"%~dp0pcf-runtime\python.exe" -B "%~dp0pcf-windows.py" %*
set "pcf_exit=%ERRORLEVEL%"
goto finish

:python_launcher
py.exe -3 -B "%~dp0pcf-windows.py" %*
set "pcf_exit=%ERRORLEVEL%"
goto finish

:python_command
python.exe -B "%~dp0pcf-windows.py" %*
set "pcf_exit=%ERRORLEVEL%"

:finish
if "%~1"=="" pause
exit /b %pcf_exit%
