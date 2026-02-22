@echo off
setlocal

echo Stopping any running taskmon instance...
call "%~dp0kill.bat"

set MSBUILD=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe
if not exist "%MSBUILD%" (
    echo ERROR: MSBuild not found.
    echo Expected: %MSBUILD%
    pause
    exit /b 1
)

echo Building taskmon.csproj ...
"%MSBUILD%" "%~dp0taskmon.csproj" /nologo /v:minimal /p:Configuration=Release

if %errorlevel% neq 0 (
    echo.
    echo Build FAILED. See errors above.
    pause
    exit /b 1
)

echo.
echo Build succeeded: %LOCALAPPDATA%\taskmon\taskmon.dll
echo You can now launch taskmon via taskmon.vbs
