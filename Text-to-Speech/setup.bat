@echo off
setlocal enabledelayedexpansion

echo Setting up Text-to-Speech application...
echo.

REM Get the directory where this script is located
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM Create a log file
set "LOG_FILE=%SCRIPT_DIR%setup_log.txt"
echo Setup started at %date% %time% > "%LOG_FILE%"

REM Check if Python is installed
echo Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo Python is not installed or not in PATH!
    echo Please install Python 3.8 or later from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    echo.
    echo After installing Python, run this setup script again.
    echo Python not found >> "%LOG_FILE%"
    pause
    exit /b 1
)

REM Remove existing virtual environment if it exists
if exist ".venv" (
    echo Removing existing virtual environment...
    rmdir /s /q ".venv"
    if errorlevel 1 (
        echo Failed to remove existing virtual environment
        echo Failed to remove .venv >> "%LOG_FILE%"
        pause
        exit /b 1
    )
)

REM Create virtual environment
echo Creating virtual environment...
python -m venv .venv
if errorlevel 1 (
    echo Failed to create virtual environment
    echo Failed to create .venv >> "%LOG_FILE%"
    pause
    exit /b 1
)

REM Activate virtual environment
echo Activating virtual environment...
call .venv\Scripts\activate
if errorlevel 1 (
    echo Failed to activate virtual environment
    echo Failed to activate .venv >> "%LOG_FILE%"
    pause
    exit /b 1
)

REM Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo Failed to upgrade pip
    echo Failed to upgrade pip >> "%LOG_FILE%"
    pause
    exit /b 1
)

REM Install setuptools and wheel
echo Installing setuptools and wheel...
pip install setuptools wheel
if errorlevel 1 (
    echo Failed to install setuptools and wheel
    echo Failed to install setuptools and wheel >> "%LOG_FILE%"
    pause
    exit /b 1
)

REM Install audio dependencies first
echo Installing audio dependencies...
pip install PyAudio
pip install sounddevice
pip install soundfile
if errorlevel 1 (
    echo Failed to install audio dependencies
    echo Failed to install audio dependencies >> "%LOG_FILE%"
    pause
    exit /b 1
)

REM Install all required packages from requirements.txt
echo Installing required packages from requirements.txt...
pip install -r requirements.txt
if errorlevel 1 (
    echo Failed to install required packages
    echo Failed to install requirements.txt packages >> "%LOG_FILE%"
    pause
    exit /b 1
)

REM Install additional audio processing dependencies
echo Installing additional audio processing dependencies...
pip install numpy
pip install scipy
pip install librosa
if errorlevel 1 (
    echo Failed to install additional audio dependencies
    echo Failed to install additional audio dependencies >> "%LOG_FILE%"
    pause
    exit /b 1
)

REM Check for FFmpeg
echo Checking for FFmpeg...
where ffmpeg >nul 2>&1
if errorlevel 1 (
    echo FFmpeg not found. Downloading and setting up FFmpeg...
    
    REM Create ffmpeg directory if it doesn't exist
    if not exist "ffmpeg" mkdir ffmpeg
    
    REM Download FFmpeg
    powershell -Command "& {Invoke-WebRequest -Uri 'https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip' -OutFile 'ffmpeg.zip'}"
    if errorlevel 1 (
        echo Failed to download FFmpeg
        echo Failed to download FFmpeg >> "%LOG_FILE%"
        pause
        exit /b 1
    )
    
    REM Extract FFmpeg
    powershell -Command "& {Expand-Archive -Path 'ffmpeg.zip' -DestinationPath 'ffmpeg' -Force}"
    if errorlevel 1 (
        echo Failed to extract FFmpeg
        echo Failed to extract FFmpeg >> "%LOG_FILE%"
        pause
        exit /b 1
    )
    
    REM Move FFmpeg files to the correct location
    move "ffmpeg\ffmpeg-master-latest-win64-gpl\bin\*" "ffmpeg\"
    rmdir /s /q "ffmpeg\ffmpeg-master-latest-win64-gpl"
    del "ffmpeg.zip"
)

REM Check for Tesseract-OCR folder
if not exist "Tesseract-OCR" (
    echo.
    echo WARNING: Tesseract-OCR folder not found!
    echo Please download Tesseract-OCR from: https://github.com/UB-Mannheim/tesseract/wiki
    echo Extract the Tesseract-OCR folder to this directory.
    echo.
    echo The application will try to use system-installed Tesseract if available,
    echo but it's recommended to have the Tesseract-OCR folder in this directory.
    echo Tesseract-OCR folder not found >> "%LOG_FILE%"
)

REM Create a configuration file for the application
echo Creating configuration file...
(
echo {
echo   "app_dir": "%SCRIPT_DIR%",
echo   "ffmpeg_path": "%SCRIPT_DIR%ffmpeg",
echo   "tesseract_path": "%SCRIPT_DIR%Tesseract-OCR",
echo   "venv_path": "%SCRIPT_DIR%.venv"
echo }
) > "app_config.json"

echo Setup completed successfully! >> "%LOG_FILE%"
echo.
echo Setup completed successfully!
echo.
echo To run the application, double-click on 'run.bat'
echo.
pause 