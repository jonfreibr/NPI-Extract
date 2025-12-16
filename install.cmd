@echo off
rem Example install batch file for Windows environment.
rem Run from the application directory
rem Required files:
rem     logo.png                    logo file displayed by application at runtime
rem     NPI Extraction.lnk          create this shortcut file to reflect your installation and runtime requirements!!
rem     npi.py                      The main application file
rem     python-3.12.3-amd64.exe     The Python source file
rem     requirements.txt            Used by Python to install required libraries
rem     unattend.xml                Required to run a "hands-off" install of Python.
rem
setlocal enabledelayedexpansion
title Install NPI Extraction Tool
echo Sourcing from: %~dp0

:: Prevent early exit on errors
set "ERRORFLAG=0"

echo Checking for existing Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python not found.
    set INSTALL=TRUE
) else (
    for /f "tokens=2 delims= " %%a in ('python --version 2^>^&1') do set "VERSION=%%a"
    echo Found Python version: !VERSION!
    if "%VERSION%" equ "3.12.3" (
        set INSTALL=FALSE
        set "PYTHON_PATH=%LocalAppData%\Programs\Python\Python312\python.exe"
    ) else (
        set INSTALL=TRUE
    )
)

if "%INSTALL%" equ "TRUE" (
    "%~dp0\python-3.12.3-amd64.exe" /passive
    set "PYTHON_PATH=%LocalAppData%\Programs\Python\Python312\python.exe"
)

set "PROJECT_DIR=%USERPROFILE%\NPIExtract"
set "VENV_DIR=%PROJECT_DIR%\venv"
set "DESKTOP_DIR=%USERPROFILE%\Desktop"

echo Creating project folder...
if not exist "%PROJECT_DIR%" md "%PROJECT_DIR%"

echo Creating virtual environment...
call "%PYTHON_PATH%" -m venv "%VENV_DIR%"
if %errorlevel% neq 0 (
    echo ERROR: Failed to create virtual environment.
    set "ERRORFLAG=1"
    goto :END
)
echo Virtual environment created at: %VENV_DIR%

echo Activating virtual environment and upgrading pip...
call "%VENV_DIR%\Scripts\activate.bat"
call %PYTHON_PATH% -m pip install --upgrade pip

if exist "%~dp0\requirements.txt" (
    echo Installing required Python packages in venv...
    call %PYTHON_PATH% -m pip install -r "%~dp0\requirements.txt"
    call %PYTHON_PATH% -m pip install -i https://PySimpleGUI.net/install PySimpleGUI
    if %errorlevel% neq 0 (
        echo ERROR: Failed to install dependencies.
        set "ERRORFLAG=1"
        goto :END
    )
) else (
    echo No requirements.txt found — skipping dependency install.
)

echo Copying files
copy /y "%~dp0\npi.py" %PROJECT_DIR%
copy /y "%~dp0\logo.png" %PROJECT_DIR%
copy /y "%~dp0\NPI Extraction.lnk" %DESKTOP_DIR%

echo Launching NPI Extraction Tool...
start "" "%DESKTOP_DIR%\NPI Extraction.lnk"
echo.

:END
if %ERRORFLAG% neq 0 (
    echo.
    echo ===============================
    echo   INSTALLATION FAILED 
    echo ===============================
    pause
) else (
    echo.
    echo ===============================
    echo   INSTALLATION COMPLETE 
    echo ===============================
    timeout /t 5
)
