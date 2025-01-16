@echo off
setlocal enabledelayedexpansion

:: Step 1: Check for Python Installation
echo Checking for Python installation...
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Python not found. Opening download page...
    start "" "https://www.python.org/downloads/"
    echo Please install Python and press Enter to continue.
    pause
    exit /b 1
) else (
    echo [SUCCESS] Python is installed.
)

:: Step 2: Check for Ruby Installation
echo Checking for Ruby installation...
where ruby >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Ruby not found. Opening download page...
    start "" "https://rubyinstaller.org/"
    echo Please install Ruby and press Enter to continue.
    pause
    exit /b 1
) else (
    echo [SUCCESS] Ruby is installed.
)

:: Step 3: Create Python Virtual Environment
echo Setting up Python virtual environment...
python -m venv venv > setup.log 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Failed to create Python virtual environment. Check setup.log for details.
    pause
    exit /b 1
) else (
    echo [SUCCESS] Python virtual environment created.
)

:: Step 4: Activate the virtual environment
if exist venv\Scripts\activate.bat (
    echo Activating Python virtual environment...
    call venv\Scripts\activate.bat
) else (
    echo [ERROR] Failed to locate virtual environment activation script.
    pause
    exit /b 1
)

:: Step 5: Install Python Dependencies
echo Installing required Python dependencies...
> requirements.txt (
    echo numpy
    echo pandas
    echo pyinstaller
    echo xlsxwriter
)

pip install -r requirements.txt >> setup.log 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Some Python dependencies failed to install. Check setup.log for details.
    pause
    exit /b 1
) else (
    echo [SUCCESS] Python dependencies installed.
)

:: Step 6: Install Required Ruby Gems if Missing
echo Checking and installing required Ruby gems...

:: Define the required gems
set gems=read crc

:: Loop through each gem
for %%G in (%gems%) do (
    echo Checking for Ruby gem: %%G...
    gem list %%G -i >nul 2>&1
    if !ERRORLEVEL! neq 0 (
        echo [INFO] Ruby gem %%G not found. Installing...
        gem install %%G --no-document >> setup.log 2>&1
        if !ERRORLEVEL! neq 0 (
            echo [ERROR] Failed to install Ruby gem %%G. Check setup.log for details.
            pause
            exit /b 1
        ) else (
            echo [SUCCESS] Ruby gem %%G installed.
        )
    ) else (
        echo [SUCCESS] Ruby gem %%G is already installed.
    )
)

echo All required Ruby gems are installed.
pause


:: Step 7: Verify Installations
echo Verifying installations...
python --version
pip list
ruby --version
gem list

:: Final Completion
echo [SUCCESS] Setup complete! Python, Ruby, and all dependencies are installed.
pause