@echo off
cd /d "%~dp0"
echo ========================================
echo Building Dryer Capacity Dashboard EXE
echo ========================================
echo.

REM Check if PyInstaller is installed
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing PyInstaller...
    pip install pyinstaller
    if %errorlevel% neq 0 (
        echo ERROR: Failed to install PyInstaller.
        pause
        exit /b 1
    )
    echo.
)

echo Running PyInstaller...
echo.

REM Include the icon and settings image in the exe bundle
pyinstaller --onefile --windowed --name "DryerCapacityDashboard" --icon=dryer.ico --add-data "dryer.ico;." --add-data "setting1.png;." CapacityDashboard.py

if %errorlevel% neq 0 (
    echo.
    echo ERROR: PyInstaller build failed!
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build complete!
echo.
echo Your executable is located at:
echo   dist\DryerCapacityDashboard.exe
echo.
echo The exe includes the icon and images - just distribute the single exe file!
echo.
echo NOTE: Self-signed certificates still show SmartScreen warnings.
echo Users click "More info" then "Run anyway" the first time.
echo ========================================
pause

echo.
echo ========================================
echo Build complete!
echo.
echo Your executable is located at:
echo   dist\DryerCapacityDashboard.exe
echo.
echo The exe includes the icon and images - just distribute the single exe file!
echo.
echo NOTE: Self-signed certificates still show SmartScreen warnings.
echo Users click "More info" then "Run anyway" the first time.
echo ========================================
pause
