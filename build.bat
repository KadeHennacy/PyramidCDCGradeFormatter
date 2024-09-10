@echo off
setlocal

:: Check if requirements.txt exists and install dependencies
if exist requirements.txt (
    echo Installing dependencies from requirements.txt...
    pip install -r requirements.txt
) else (
    echo requirements.txt not found. Continuing without installing dependencies...
)

:: Check if PyInstaller is installed
pip show pyinstaller >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo PyInstaller is not installed. Installing...
    pip install pyinstaller
)

:: Package the script with PyInstaller
pyinstaller --onefile --clean main.py

echo Build completed.
endlocal
