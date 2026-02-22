@echo off
call "%~dp0build.bat"
if %errorlevel% neq 0 (
    echo Build failed — taskmon not launched.
    pause
    exit /b 1
)

echo.
echo Launching taskmon...
start "" wscript.exe "%~dp0taskmon.vbs"
