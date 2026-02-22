@echo off
setlocal enabledelayedexpansion

if "%~1"=="" (
    echo Usage: removebg ^<image_file^>
    echo Example: removebg c:/images/photo.jpg
    exit /b 1
)

set "IMAGE_FILE=%~1"

if not exist "!IMAGE_FILE!" (
    echo Error: File not found: !IMAGE_FILE!
    exit /b 1
)

set "OUTPUT_FILE=%~dpn1_nobg%~x1"

echo Removing background from image...
echo Input:  !IMAGE_FILE!
echo Output: !OUTPUT_FILE!
echo.

rembg i -m birefnet-portrait "!IMAGE_FILE!" "!OUTPUT_FILE!"

if errorlevel 1 (
    echo Error: Background removal failed
    exit /b 1
)

echo.
echo Background removal complete!
echo Output saved to: !OUTPUT_FILE!
