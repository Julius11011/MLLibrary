@echo off
setlocal enabledelayedexpansion
title MLC Audio Podcast Cloud Uploader

cd /d "%~dp0"

echo ===============================================================================
echo                MLC AUDIO PODCAST GITHUB CLOUD UPLOADER
echo ===============================================================================
echo.
echo This tool manages high-speed CDN cloud audio streaming for your study materials.
echo.

python "%~dp0upload_audio_to_github_release.py"

echo.
echo ===============================================================================
echo Tip: If you have a GitHub Personal Access Token, you can run:
echo      python upload_audio_to_github_release.py --token YOUR_TOKEN
echo ===============================================================================
echo.
pause
