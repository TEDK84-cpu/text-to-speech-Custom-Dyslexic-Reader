@echo off
echo Setting up Text-to-Speech application...
echo.

REM Get the directory where this script is located
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Python is not installed or not in PATH!
    echo Please install Python 3.8 or later from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    echo.
    echo After installing Python, run this setup script again.
    pause
    exit /b 1
)

REM Create virtual environment if it doesn't exist
if not exist .venv (
    echo Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo Failed to create virtual environment
        echo Please make sure Python is installed and in your PATH
        pause
        exit /b 1
    )
)

REM Activate virtual environment
call .venv\Scripts\activate
if errorlevel 1 (
    echo Failed to activate virtual environment
    pause
    exit /b 1
)

REM Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo Failed to upgrade pip
    pause
    exit /b 1
)

REM Install setuptools and wheel
echo Installing setuptools and wheel...
pip install setuptools wheel
if errorlevel 1 (
    echo Failed to install setuptools and wheel
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
    pause
    exit /b 1
)

REM Install all required packages from requirements.txt
echo Installing required packages from requirements.txt...
pip install -r requirements.txt
if errorlevel 1 (
    echo Failed to install required packages
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
        pause
        exit /b 1
    )
    
    REM Extract FFmpeg
    powershell -Command "& {Expand-Archive -Path 'ffmpeg.zip' -DestinationPath 'ffmpeg' -Force}"
    if errorlevel 1 (
        echo Failed to extract FFmpeg
        pause
        exit /b 1
    )
    
    REM Move FFmpeg files to the correct location
    move "ffmpeg\ffmpeg-master-latest-win64-gpl\bin\*" "ffmpeg\"
    rmdir /s /q "ffmpeg\ffmpeg-master-latest-win64-gpl"
    del "ffmpeg.zip"
)

REM Set FFmpeg path for this session and create a temporary environment variable
set "FFMPEG_PATH=%SCRIPT_DIR%\ffmpeg"
set "PATH=%PATH%;%FFMPEG_PATH%"

REM Create a temporary batch file to set FFmpeg path
echo @echo off > set_ffmpeg_path.bat
echo set "PATH=%%PATH%%;%FFMPEG_PATH%" >> set_ffmpeg_path.bat
echo set "FFMPEG_PATH=%FFMPEG_PATH%" >> set_ffmpeg_path.bat
echo start "" "run.bat" >> set_ffmpeg_path.bat

REM Check for Tesseract-OCR folder
if not exist "Tesseract-OCR" (
    echo.
    echo WARNING: Tesseract-OCR folder not found!
    echo Please download Tesseract-OCR from: https://github.com/UB-Mannheim/tesseract/wiki
    echo Extract the Tesseract-OCR folder to this directory.
    echo.
    echo The application will try to use system-installed Tesseract if available,
    echo but it's recommended to have the Tesseract-OCR folder in this directory.
)

echo.
echo Setup completed successfully!
echo.
echo To run the application, double-click on 'run.bat'
echo.
pause 