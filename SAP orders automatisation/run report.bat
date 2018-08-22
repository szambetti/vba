@echo off
TITLE Daily orders generator
SETLOCAL ENABLEEXTENSIONS

if exist template.xlsb (
    echo Starting the report generator...
	start template.xlsb /popup
    echo.
    echo This window will automatically close in 10 seconds.. 
) else (
echo Error: Template file does not exist, please check the instructions...
)
timeout 10 >nul
exit
