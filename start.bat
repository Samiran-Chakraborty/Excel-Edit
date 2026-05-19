@echo off
title Excel Editor - Starting Servers

echo.
echo  ====================================
echo   Excel Editor - Starting Servers
echo  ====================================
echo.

echo  [1/2] Starting Backend (port 3000)...
start "Excel Backend" cmd /k "cd /d "%~dp0backend" && node server.js"

echo  [2/2] Starting Frontend (port 5173)...
start "Excel Frontend" cmd /k "cd /d "%~dp0frontend" && npm run dev"

ping -n 5 127.0.0.1 >nul

echo.
echo  Opening app in browser...
start "" "http://localhost:5173"
