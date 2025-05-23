@echo off
echo Starting Text-to-Speech application...

REM Get the directory where this script is located
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM Set FFmpeg path
set "FFMPEG_PATH=%SCRIPT_DIR%\ffmpeg"
set "PATH=%PATH%;%FFMPEG_PATH%"

REM Activate virtual environment
call .venv\Scripts\activate
if errorlevel 1 (
    echo Failed to activate virtual environment
    echo Please run setup.bat first
    pause
    exit /b 1
)

REM Verify FFmpeg is accessible
where ffmpeg >nul 2>&1
if errorlevel 1 (
    echo FFmpeg not found in PATH
    echo Current PATH: %PATH%
    echo FFmpeg should be in: %FFMPEG_PATH%
    pause
    exit /b 1
)

echo Running application...
python Text-to-Speech.py
if errorlevel 1 (
    echo Application failed to start
    pause
    exit /b 1
)

pause 