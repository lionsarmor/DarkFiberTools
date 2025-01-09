@echo off
setlocal enabledelayedexpansion

:: Step 1: Check for Python installation
echo Checking Python installation...
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Python not found. Installing Python using winget...
    winget install -e --id Python.Python.3 -s winget
    if %ERRORLEVEL% neq 0 (
        echo Failed to install Python. Ensure winget is functional and you have internet access.
        pause
        exit /b 1
    )
    echo Python installed successfully.
) else (
    echo Python is already installed.
)

:: Step 2: Check for Ruby installation
echo Checking Ruby installation...
where ruby >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Ruby not found. Installing Ruby using winget...
    winget install -e --id RubyInstallerTeam.Ruby -s winget
    if %ERRORLEVEL% neq 0 (
        echo Failed to install Ruby. Ensure winget is functional and you have internet access.
        pause
        exit /b 1
    )
    echo Ruby installed successfully.
) else (
    echo Ruby is already installed.
)

:: Step 3: Ensure Ruby Gems are Installed
echo Checking Ruby gems...

:: Install 'read' gem
gem list -i read >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Installing 'read' gem...
    gem install read --no-document
    if %ERRORLEVEL% neq 0 (
        echo Failed to install 'read' gem. Ensure RubyGems is functional and internet is available.
        pause
        exit /b 1
    )
) else (
    echo 'read' gem is already installed.
)

:: Install 'crc' gem
gem list -i crc >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Installing 'crc' gem...
    gem install crc --no-document
    if %ERRORLEVEL% neq 0 (
        echo Failed to install 'crc' gem. Ensure RubyGems is functional and internet is available.
        pause
        exit /b 1
    )
) else (
    echo 'crc' gem is already installed.
)

:: Step 4: Create Python Virtual Environment
echo Setting up Python virtual environment...
if not exist venv (
    python -m venv venv
    if %ERRORLEVEL% neq 0 (
        echo Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo Virtual environment created.
)

call venv\Scripts\activate

:: Step 5: Upgrade pip, setuptools, and wheel
echo Upgrading pip, setuptools, and wheel...
pip install --upgrade pip setuptools wheel >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Failed to upgrade pip, setuptools, and wheel.
    pause
    exit /b 1
)

:: Step 6: Install Required Python Dependencies
echo Installing required Python dependencies...
(
    echo numpy
    echo pandas
    echo pyinstaller
    echo xlsxwriter
) > requirements.txt

pip install -r numpy
pip install -r pandas
pip install -r xlsxwriter
pip install -r requirements.txt
if %ERRORLEVEL% neq 0 (
    echo Failed to install Python dependencies. Please check for errors in your Python environment.
    pause
    exit /b 1
)

:: Step 7: Final Message
echo Environment setup complete! Python, Ruby, and dependencies are installed.
pause
