@echo off
setlocal enabledelayedexpansion
title MLC Cloud Sync & Auto-Deploy

cd /d "%~dp0"

echo ===============================================================================
echo                   MLC CLOUD SYNC & AUTO-DEPLOYMENT
echo ===============================================================================
echo.

echo [1/4] Scanning and regenerating all TTS HTML readers and Study Hub...
python "%~dp0generate_tts_reader.py"

echo.
echo [2/4] Checking Git status and staging updates...
git add .
git commit -m "Update MLC legal notes, HTML readers, and study hub %date% %time%"
git branch -M main
git push origin main

echo.
echo [3/4] Deploying updated web assets to Cloudflare Workers...
call npx wrangler deploy

echo.
echo [4/4] Checking audio podcast sync...
echo If you generated new .mp3 podcast files, you can upload them to the GitHub CDN
echo by running 'Upload-Audio-To-GitHub.bat' or python upload_audio_to_github_release.py.
echo.
echo ===============================================================================
echo [SUCCESS] Everything is synchronized and deployed!
echo Live Site: https://mllibrary.juliusrayn-balitbit.workers.dev
echo ===============================================================================
echo.
pause
