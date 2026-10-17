@echo off
rem Table OCR - starts the local server and opens the tool in your browser.
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo First run: creating Python environment...
  python -m venv .venv || goto :err
  ".venv\Scripts\python.exe" -m pip install pymupdf openpyxl || goto :err
)
".venv\Scripts\python.exe" server.py %*
pause
exit /b
:err
echo Setup failed. Make sure Python 3 is installed and on PATH.
pause
