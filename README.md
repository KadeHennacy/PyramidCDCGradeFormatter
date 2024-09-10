# README for Spreadsheet Formatter Application

## Project Overview

This Python application is designed to format spreadsheets exported from PyramidCDC's learning platforms. It offers customization options such as resizing columns, centering text, and applying word wraps, enhancing readability and presentation of spreadsheet data. The application is built using `Tkinter` for the GUI, making it user-friendly.

## User Guide

### System Requirements

- This application is designed primarily for Windows. To use it on macOS, you will need to replace the Windows-specific batch file (`build.bat`) with a shell script equivalent for macOS.

### Installing Python

- Download and install the latest version of Python from the [official Python website](https://www.python.org/downloads/). Ensure to check the option 'Add Python to PATH' at the beginning of the installation process.

### Running the Application on Windows

1. **Download the Project**: Clone or download the project folder to your local machine.

```
git clone https://github.com/KadeHennacy/PyramidCDCGradeFormatter.git
```

2. **Open Command Prompt**: Navigate to the directory where the project is located using the Command Prompt.

```
cd C:\Users\kadeh\Projects\Force For Good\PyramidCDCGradeFormatter
```

3. **Build the Executable**:
   - Type `./build.bat` in the command prompt and press Enter. This script automates the setup process, including installing necessary dependencies and packaging the application into a single executable file using PyInstaller.
4. **Run the Application**:
   - Navigate to the `dist` folder created by PyInstaller, and double-click on the executable file to run the application.
   - If you encounter a security alert saying "Windows protected your PC", click on "More info" and then select "Run anyway" to start the application.

### Developer Guide

#### Dependencies

- The application uses the following Python packages:
  - `pandas` for data manipulation.
  - `openpyxl` for reading and writing Excel files.
  - `tkinter` for the graphical user interface.

#### Installing Dependencies

- Dependencies are listed in the `requirements.txt` file. To install these, navigate to the project directory in your command prompt or terminal and run:
  ```
  pip install -r requirements.txt
  ```

#### About `build.bat`

- The `build.bat` file is a script for Windows that prepares the application for distribution:
  - It checks for and installs any missing dependencies from `requirements.txt`.
  - It verifies whether PyInstaller is installed, installs it if it isn't, and then uses it to package the application into a standalone executable.
  - The `--clean` flag in the PyInstaller command ensures that any previous builds are cleared before creating a new one, preventing issues from outdated files.
  - The `--onefile` flag bundles everything into a single executable making it easier to execute and distribute

#### Understanding the Script

- The script allows users to load a CSV or Excel file, process the data based on the selected settings, and then save the formatted output as an Excel file.
- It includes detailed comments explaining how it works.
- You can edit the script with any text editor like VSCode and running `python main.py`

## Contribute

- To contribute to this project, you can fork the repository, make changes, and submit a pull request. Contributions to enhance features, improve the user interface, or fix bugs are welcome.
