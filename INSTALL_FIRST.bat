@echo off
cd /d "%~dp0"
where py >nul 2>&1
if %errorlevel%==0 (set "PYTHON=py -3") else (set "PYTHON=python")
%PYTHON% -m venv .venv
if errorlevel 1 (echo Python 3 must be installed. & pause & exit /b 1)
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (echo Installation failed. & pause & exit /b 1)
echo Installation completed. You can now run RUN_REPORT.bat
pause
