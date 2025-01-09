@echo off
setlocal enabledelayedexpansion

:: --- Logging Prefix ---
set "LOG_PREFIX=[SetupScript]"
set "LOG_INFO=[INFO] "
set "LOG_ERROR=[ERROR]"

:: --- Step 1: Check for Python installation ---
echo %LOG_INFO% Checking for Python installation...
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo %LOG_ERROR% Python not found. Installing Python using winget...
    winget install -e --id Python.Python.3 -s winget
    if %ERRORLEVEL% neq 0 (
        echo %LOG_ERROR% Failed to install Python. Ensure winget is functional and you have internet access.
        pause
        exit /b 1
    )
    echo %LOG_INFO% Python installed successfully.
) else (
    echo %LOG_INFO% Python is already installed.
)

:: --- Step 2: Check for Ruby installation ---
echo %LOG_INFO% Checking for Ruby installation...
where ruby >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo %LOG_ERROR% Ruby not found. Installing Ruby using winget...
    winget install -e --id RubyInstallerTeam.Ruby -s winget
    if %ERRORLEVEL% neq 0 (
        echo %LOG_ERROR% Failed to install Ruby. Ensure winget is functional and you have internet access.
        pause
        exit /b 1
    )
    echo %LOG_INFO% Ruby installed successfully.
) else (
    echo %LOG_INFO% Ruby is already installed.
)

:: --- Step 3: Check and Install Required Ruby Gems ---
echo %LOG_INFO% Checking for required Ruby gems...
for %%G in (read crc) do (
    echo %LOG_INFO% Checking for Ruby gem: %%G...
    gem list -i %%G >nul 2>&1
    if !ERRORLEVEL! neq 0 (
        echo %LOG_INFO% Installing Ruby gem: %%G...
        gem install %%G --no-document >gem_install_log.txt 2>&1
        if !ERRORLEVEL! neq 0 (
            echo %LOG_ERROR% Failed to install Ruby gem %%G. Check 'gem_install_log.txt' for details.
            type gem_install_log.txt
            pause
            exit /b 1
        ) else (
            echo %LOG_INFO% Ruby gem %%G installed successfully.
        )
    ) else (
        echo %LOG_INFO% Ruby gem %%G is already installed.
    )
)

:: --- Step 4: Create Python Virtual Environment ---
echo %LOG_INFO% Setting up Python virtual environment...
if not exist venv (
    python -m venv venv
    if %ERRORLEVEL% neq 0 (
        echo %LOG_ERROR% Failed to create Python virtual environment.
        pause
        exit /b 1
    )
    echo %LOG_INFO% Virtual environment created.
) else (
    echo %LOG_INFO% Virtual environment already exists.
)

:: Activate the virtual environment
call venv\Scripts\activate
if %ERRORLEVEL% neq 0 (
    echo %LOG_ERROR% Failed to activate the Python virtual environment.
    pause
    exit /b 1
)

:: --- Step 5: Upgrade pip, setuptools, and wheel ---
echo %LOG_INFO% Upgrading pip, setuptools, and wheel...
venv\Scripts\python.exe -m pip install --upgrade pip setuptools wheel >pip_upgrade_log.txt 2>&1
if %ERRORLEVEL% neq 0 (
    echo %LOG_ERROR% Failed to upgrade pip, setuptools, and wheel.
    echo Check 'pip_upgrade_log.txt' for details.
    type pip_upgrade_log.txt
    pause
    exit /b 1
) else (
    echo %LOG_INFO% pip, setuptools, and wheel upgraded successfully.
)

:: --- Step 6: Install Python Dependencies ---
echo %LOG_INFO% Installing required Python dependencies...
(
    echo numpy
    echo pandas
    echo pyinstaller
    echo xlsxwriter
) > requirements.txt

pip install -r requirements.txt >pip_install_log.txt 2>&1
if %ERRORLEVEL% neq 0 (
    echo %LOG_ERROR% Some Python dependencies failed to install. Check 'pip_install_log.txt' for details.
    type pip_install_log.txt
    pause
    exit /b 1
) else (
    echo %LOG_INFO% All Python dependencies installed successfully.
)

:: --- Step 7: Verify Python Dependencies ---
echo %LOG_INFO% Verifying Python dependencies...
for %%P in (numpy pandas pyinstaller xlsxwriter) do (
    pip show %%P >nul 2>&1
    if !ERRORLEVEL! neq 0 (
        echo %LOG_ERROR% Dependency %%P is missing. Attempting to reinstall...
        pip install %%P >nul 2>&1
        if !ERRORLEVEL! neq 0 (
            echo %LOG_ERROR% Failed to install %%P. Check 'pip_install_log.txt' for details.
            pause
            exit /b 1
        ) else (
            echo %LOG_INFO% Dependency %%P installed successfully.
        )
    ) else (
        echo %LOG_INFO% Dependency %%P is properly installed.
    )
)

:: --- Completion ---
echo %LOG_INFO% Environment setup complete! Python, Ruby, and all dependencies are installed and verified.
pause
