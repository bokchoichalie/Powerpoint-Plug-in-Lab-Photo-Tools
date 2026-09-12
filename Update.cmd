@echo off
setlocal
title Lab Photo Tools 0.1.11 - Update
echo Save your files and close all PowerPoint windows before updating.
echo Existing Python runtime and background model will be reused.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Install.ps1" -SourceDirectory "%~dp0."
set "LAB_UPDATE_RESULT=%ERRORLEVEL%"
if "%LAB_UPDATE_RESULT%"=="0" (
    echo Update complete. Reopen PowerPoint to use Lab Photo Tools 0.1.11.
) else (
    echo Update failed. Read the error above. Close PowerPoint and try again.
)
pause
exit /b %LAB_UPDATE_RESULT%

