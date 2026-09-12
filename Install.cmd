@echo off
setlocal
title Lab Photo Tools - Install
echo Lab Photo Tools preview - current-user installation
echo Close PowerPoint before continuing. Python 3.12 x64 is required.
echo This setup downloads dependencies and a model once. Photos stay on this PC.
echo.
set "LAB_SETUP_MODE=-PrepareRuntime"
if /i "%~1"=="/check" set "LAB_SETUP_MODE=-CheckOnly"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Install.ps1" -SourceDirectory "%~dp0." %LAB_SETUP_MODE%
set "LAB_RESULT=%ERRORLEVEL%"
echo.
if "%LAB_RESULT%"=="0" (
    if /i "%~1"=="/check" (
        echo Source path check completed. Nothing was installed.
    ) else (
        echo Installation completed. Open PowerPoint and find the Lab Photo Tools tab.
    )
) else (
    echo Installation failed. Please read the error above and see README.md.
)
if /i not "%~1"=="/check" pause
exit /b %LAB_RESULT%

