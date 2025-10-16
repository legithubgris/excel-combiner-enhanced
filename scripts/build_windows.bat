@echo off
REM Build script for Windows
REM This script creates a standalone Windows executable

echo Building Excel Combiner for Windows...

REM Check if required packages are installed
echo Installing required packages...
pip install pandas openpyxl xlrd pyinstaller

REM Create the Windows executable
echo Creating Windows executable...
pyinstaller excel_combiner_windows.spec --clean --noconfirm

REM Check if build was successful
if exist "dist\ExcelCombiner.exe" (
    echo ✅ Windows build completed successfully!
    echo 📁 Executable created at: dist\ExcelCombiner.exe
    echo.
    echo To run the application:
    echo   Double-click on ExcelCombiner.exe in the dist folder
    echo   Or run: dist\ExcelCombiner.exe
    echo.
    echo To create a distributable package:
    echo   You can compress the dist folder to create a .zip file
    echo   Or use tools like Inno Setup to create an installer
) else (
    echo ❌ Build failed. Check the output above for errors.
    pause
    exit /b 1
)

pause