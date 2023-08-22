@echo off
setlocal enabledelayedexpansion

REM Set the terminal title
title (wagt) Row2Chat

REM Set the path to your virtual environment's activate script
set "venv_activate=%USERPROFILE%\Desktop\wagt\Scripts\activate"

REM Keep looping
:loop

REM Activate the virtual environment
call "%venv_activate%"

REM Clear the screen
cls

color A

REM Display instructions for users
echo This script allows you to run a wagt script with specific arguments.
echo Please provide the necessary arguments.
echo Type -h or --help for help menu.
echo.

REM Prompt the user for arguments
set /p "python_arguments=(wagt) arguments: "

REM Check if the user wants to exit
if /i "!python_arguments!"=="exit" (
    exit /b
)

REM Run the Python file "main.py" with provided arguments
set "python_command=python main.py !python_arguments!"
!python_command!

REM Pause to keep the terminal window open
pause

goto loop
