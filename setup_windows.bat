@echo off
setlocal
cd /d "%~dp0"
py -3.12 --version >nul 2>&1
if errorlevel 1 (
  echo Install 64-bit Python 3.12 from python.org with the Python launcher enabled.
  pause
  exit /b 1
)
if not exist ".venv\Scripts\python.exe" (
  py -3.12 -m venv .venv
  if errorlevel 1 goto fail
)
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto fail
echo Setup complete. Double-click run_windows.bat to open the app.
pause
exit /b 0
:fail
echo Setup failed. Read the message above; check your internet connection and Python installation.
pause
exit /b 1
