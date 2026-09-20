@echo off
title MLC Local Study Server
cd /d "%~dp0"
echo ===============================================================================
echo                MLC LOCAL STUDY SERVER (PORT 8080)
echo ===============================================================================
echo.
echo Serving study materials at:
echo   Hub Index: http://localhost:8080/MLC_Study_Hub.html
echo.
echo Press Ctrl+C in this window to stop the server anytime.
echo.
start http://localhost:8080/MLC_Study_Hub.html
python -m http.server 8080
