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
".venv\Scripts\python.exe" -m pip install -r requirements-build.txt
if errorlevel 1 goto fail
".venv\Scripts\python.exe" -m unittest discover -s tests -v
if errorlevel 1 goto fail
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onefile --windowed --name EDocGenerator --hidden-import pythoncom --hidden-import pywintypes --collect-submodules win32com launcher.py
if errorlevel 1 goto fail
echo Built: %cd%\dist\EDocGenerator.exe
echo Desktop Microsoft Excel must be installed on the computer running the exe.
pause
exit /b 0
:fail
echo Build failed. Read the message above.
pause
exit /b 1
