@echo off
setlocal enabledelayedexpansion
title MLC Text-to-Speech HTML Reader Generator

cd /d "%~dp0"

echo ===============================================================================
echo                MLC TEXT-TO-SPEECH (READ ALOUD) HTML GENERATOR
echo ===============================================================================
echo.

if not "%~1"=="" (
    echo [INFO] Direct file/folder dropped onto script: "%~1"
    python "%~dp0generate_tts_reader.py" "%~1"
    echo.
    echo [DONE] Converted "%~1" to HTML Reader.
    pause
    exit /b
)

:MENU
echo Please select an option:
echo.
echo  [1] Convert ALL study notes across all MLC subfolders (.md, .docx, .txt, .pdf)
echo  [2] Convert Criminal Law folder
echo  [3] Convert Constitutional Law (CsL) folder
echo  [4] Convert Statutory Construction folder
echo  [5] Convert Criminal Procedure folder
echo  [6] Convert BLJE / Legal Ethics folder
echo  [7] Refresh Master Study Hub Index (MLC_Study_Hub.html & index.html)
echo  [8] Start Local Web Server (http://localhost:8080)
echo  [9] Sync & Auto-Deploy to Cloud (GitHub / Cloudflare)
echo  [0] Exit
echo.
set /p opt="Enter choice (0-9): "

if "%opt%"=="1" (
    echo.
    echo Running full scan across all MLC subfolders...
    python "%~dp0generate_tts_reader.py" --all
    echo.
    echo Master Study Hub generated at: %~dp0MLC_Study_Hub.html
    pause
    goto MENU
)

if "%opt%"=="2" (
    echo.
    python "%~dp0generate_tts_reader.py" --dir "%~dp0First Sem 1st Year\Criminal Law"
    pause
    goto MENU
)

if "%opt%"=="3" (
    echo.
    python "%~dp0generate_tts_reader.py" --dir "%~dp0First Sem 1st Year\CsL"
    pause
    goto MENU
)

if "%opt%"=="4" (
    echo.
    python "%~dp0generate_tts_reader.py" --dir "%~dp0First Sem 1st Year\Statutory Construction"
    pause
    goto MENU
)

if "%opt%"=="5" (
    echo.
    python "%~dp0generate_tts_reader.py" --dir "%~dp0First Sem 1st Year\Criminal Procedure"
    pause
    goto MENU
)

if "%opt%"=="6" (
    echo.
    python "%~dp0generate_tts_reader.py" --dir "%~dp0First Sem 1st Year\BLJE"
    pause
    goto MENU
)

if "%opt%"=="7" (
    echo.
    python "%~dp0generate_tts_reader.py" --hub
    pause
    goto MENU
)

if "%opt%"=="8" (
    echo.
    echo Starting local web server on port 8080...
    start "" cmd /k "cd /d "%~dp0" && python -m http.server 8080"
    start http://localhost:8080/MLC_Study_Hub.html
    goto MENU
)

if "%opt%"=="9" (
    call "%~dp0update_and_sync_cloud.bat"
    goto MENU
)

if "%opt%"=="0" (
    exit /b
)

echo Invalid choice. Please try again.
echo.
goto MENU
