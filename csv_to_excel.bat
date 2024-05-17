@echo off
REM Get the directory of the current script
set SCRIPT_DIR=%~dp0

REM Define the Python script name
set PYTHON_SCRIPT=csv_to_excel.py

REM Execute the Python script
python "%SCRIPT_DIR%%PYTHON_SCRIPT%"
