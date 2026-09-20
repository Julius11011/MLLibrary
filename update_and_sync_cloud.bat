@echo off
setlocal enabledelayedexpansion
title MLC Cloud Sync & Auto-Deploy

cd /d "%~dp0"

echo ===============================================================================
echo                   MLC CLOUD SYNC & AUTO-DEPLOYMENT
echo ===============================================================================
echo.

echo [1/3] Scanning and regenerating all TTS HTML readers and Study Hub...
python "%~dp0generate_tts_reader.py" --all

echo.
echo [2/3] Checking Git status...

if not exist "%~dp0.git" (
    echo.
    echo [INFO] Git repository is not initialized yet.
    echo Initializing local repository on branch 'main'...
    git init -b main
    echo.
    echo Please create a new repository on GitHub (e.g., 'mlc-law-library')
    set /p remote_url="Enter your GitHub Repository URL (https://github.com/USERNAME/REPO.git): "
    if not "!remote_url!"=="" (
        git remote add origin !remote_url!
    )
)

echo.
set msg=Update legal notes and TTS readers %date% %time%
set /p user_msg="Enter commit note (or press Enter for default): "
if not "!user_msg!"=="" set msg=!user_msg!

echo.
echo [3/3] Staging, committing, and pushing to cloud...
git add .
git commit -m "%msg%"
git branch -M main
git push -u origin main

if %errorlevel% equ 0 (
    echo.
    echo ===============================================================================
    echo [SUCCESS] Pushed to GitHub!
    echo Your online website (GitHub Pages / Cloudflare Pages) will update in ~30s!
    echo ===============================================================================
) else (
    echo.
    echo [NOTE] Push encountered an issue. Please verify your GitHub credentials or remote URL.
)

echo.
pause
