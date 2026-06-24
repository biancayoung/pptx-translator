@echo off
chcp 65001 >nul
set PYTHONUTF8=1
title PPTX Translator App
echo =========================================
echo         PPTX TRANSLATOR APP
echo =========================================
echo.

:: Check if a file was dragged and dropped onto the app
set "input_file=%~1"

:: If not, ask the user to type or paste the path
if "%input_file%"=="" (
    echo [TIP] You can drag and drop a .pptx file directly onto this window!
    echo.
    set /p "input_file=Please paste the path to your .pptx file and press Enter: "
)

:: Remove surrounding quotes if they exist
set "input_file=%input_file:"=%"

echo.
echo Translating: "%input_file%"
echo Please wait...
echo.

:: Navigate to the folder where this batch script is located
cd /d "%~dp0"

:: Run the python script
python pptx_translator.py -i "%input_file%"

echo.
echo =========================================
echo Finished! Check the folder for your translated file.
pause
