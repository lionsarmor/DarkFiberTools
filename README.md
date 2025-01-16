
# DarkFiberTools Setup and Usage Guide

This guide will help you set up the DarkFiberTools application on Windows, macOS, and Linux. It covers installation steps, setting up the virtual environment, installing dependencies, and launching the application.

## Windows Instructions

### Step 1: Install Dependencies

1. **Install Python**  
   If you don't have Python installed, download and install it from the official Python website: [https://www.python.org/downloads/](https://www.python.org/downloads/).

   Ensure to add Python to your PATH during installation.

2. **Install Ruby**  
   To install Ruby, follow these steps:
   - Download Ruby from [https://rubyinstaller.org/](https://rubyinstaller.org/).
   - Run the installer and ensure that Ruby is added to your system PATH.
   - Optionally, you can install RubyGems if it’s not bundled with the installer.

3. **Run the Setup Script**
   - Right-click on `setup.bat` and select **Run as administrator**. This will attempt to install all necessary dependencies for the application.
   - If the console says failed, open the setup.log file, if the log says everything installed successfully move on to step 2. if there were failures, install the dependencies      manually, see trouble shooting for windows.

### Step 2: Launch the Application

1. **Launch the Program**
   - Double-click on `launch.bat` to run the application. It will activate the virtual environment (if it exists) and then execute `analyzer.py`.

2. **Troubleshooting**
   - If you face any errors, they will appear in the console. Make sure that all dependencies are installed properly, and check the steps below if you need to manually install them.

---

## macOS Instructions

### Step 1: Install Dependencies

1. **Install Python**  
   If you don't have Python installed, open a terminal and install it via Homebrew:
   ```bash
   brew install python
   ```

2. **Install Ruby**  
   To install Ruby, run:
   ```bash
   brew install ruby
   ```

3. **Install Dependencies Manually**
   - After installing Python and Ruby, you need to install the Python dependencies using `pip`:
   ```bash
   pip install -r requirements.txt
   ```

### Step 2: Launch the Application

1. **Running the Program**
   - Open a terminal, navigate to the application directory, and run the following command to start the program:
   ```bash
   python analyzer.py
   ```

2. **Troubleshooting**
   - If dependencies are not installed correctly, run the following to install them manually:
   ```bash
   pip install numpy pandas pyinstaller xlsxwriter
   ```

---

## Linux Instructions

### Step 1: Install Dependencies

1. **Install Python**  
   To install Python, use your package manager:
   ```bash
   sudo apt update
   sudo apt install python3 python3-pip
   ```

2. **Install Ruby**  
   Install Ruby using the following command:
   ```bash
   sudo apt install ruby
   ```

3. **Install Dependencies**
   - Once Python and Ruby are installed, run:
   ```bash
   pip3 install -r requirements.txt
   ```

### Step 2: Launch the Application

1. **Running the Program**
   - Open a terminal, navigate to the application directory, and run the following:
   ```bash
   python3 analyzer.py
   ```

2. **Troubleshooting**
   - If dependencies are missing, run:
   ```bash
   pip3 install numpy pandas pyinstaller xlsxwriter
   ```

---

## Creating a Desktop Shortcut

### Windows

1. Right-click on your Desktop and select **New > Shortcut**.
2. In the location field, type the following:
   ```bash
   "C:\path\to\your\project\launch.bat"
   ```
3. Click **Next**, give the shortcut a name (e.g., "DarkFiberTools"), and click **Finish**.

### macOS and Linux

For macOS and Linux, you can create a desktop shortcut by creating a `.desktop` file.

1. Open a text editor and create a new file with the following content:
   ```bash
   [Desktop Entry]
   Version=1.0
   Name=DarkFiberTools
   Exec=python3 /path/to/your/project/analyzer.py
   Icon=/path/to/icon.png
   Terminal=true
   Type=Application
   Categories=Utility;
   ```

2. Save this file with a `.desktop` extension (e.g., `DarkFiberTools.desktop`).

3. Make it executable by running:
   ```bash
   chmod +x /path/to/DarkFiberTools.desktop
   ```

4. Move the `.desktop` file to your Desktop directory:
   ```bash
   mv /path/to/DarkFiberTools.desktop ~/Desktop/
   ```

---

## Troubleshooting

### Common Issues

- **Virtual Environment Not Found**:  
  If you are on Windows and the virtual environment does not exist, you can manually create it by running:
  ```bash
  python -m venv venv
  ```

  Then activate it by running:
  ```bash
  .\venv\Scripts\activate
  ```

- **Dependency Installation Issues**:  
  If the dependencies fail to install during setup, try manually installing them by running:
  ```bash
  pip install -r requirements.txt
  ```

---

That should cover the setup and usage of the application on Windows, macOS, and Linux systems! If you run into any issues or need further assistance, feel free to reach out.
