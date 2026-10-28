@echo off
rem Extracts the land table from the award in every folder under the M3 Data folder (or the folder
rem given as the first argument) and saves "Award <no> <mouza> - Land.xlsx" next to each award.
rem Add --redo to re-extract folders that already have their Excel file.
cd /d "%~dp0"
if "%~1"=="" (
    ".venv\Scripts\python.exe" award_batch.py "D:\Danish Enterprises\Motorway\M3 Data"
) else (
    ".venv\Scripts\python.exe" award_batch.py %*
)
pause
