@echo off
setlocal
title Lab Photo Tools - Uninstall
echo Lab Photo Tools - remove the current-user installation
echo Close PowerPoint first. Presentation files will be kept.
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Uninstall.ps1"
set "LAB_RESULT=%ERRORLEVEL%"
echo.
if "%LAB_RESULT%"=="0" (
    echo Lab Photo Tools was removed.
) else (
    echo Uninstallation failed. Please read the error above and see README.md.
)
pause
exit /b %LAB_RESULT%

