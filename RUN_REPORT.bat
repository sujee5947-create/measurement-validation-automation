@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Please run INSTALL_FIRST.bat first.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" main.py
if errorlevel 1 (
  echo Analysis failed. Check the error above.
  pause
  exit /b 1
)
echo.
echo DONE! Updated PPTX / HTML / XLSX files are in outputs folder.
pause
