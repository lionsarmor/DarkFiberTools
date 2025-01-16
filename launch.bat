@echo off
setlocal

:: Check if virtual environment exists
if exist venv (
    echo Activating virtual environment...
    call venv\Scripts\activate
) else (
    echo No virtual environment found. Running directly...
)

:: Run analyzer.py
echo Launching analyzer.py...
python analyzer.py

:: Pause to view any errors or messages
pause
