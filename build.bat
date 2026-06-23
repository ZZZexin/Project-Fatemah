@echo off
:: ====== EDIT THIS LINE TO BUMP THE VERSION =================================
set VERSION=0.1.2
:: ==========================================================================
set APP_NAME=AutoTVC_v%VERSION%

echo ============================================================
echo  Building %APP_NAME%
echo ============================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: python not found on PATH.
    pause & exit /b 1
)

echo Installing build + runtime dependencies...
python -m pip install pyinstaller pywinauto psutil pyperclip comtypes --quiet --upgrade

echo.
echo Building exe (1-3 minutes the first time)...
python -m PyInstaller build.spec --noconfirm --clean

if errorlevel 1 (
    echo.
    echo BUILD FAILED -- see output above.
    pause & exit /b 1
)

echo.
echo Zipping...
powershell -Command "Compress-Archive -Path 'dist\%APP_NAME%.exe' -DestinationPath 'dist\%APP_NAME%.zip' -Force"

echo.
echo ============================================================
echo  Done!
echo    dist\%APP_NAME%.exe  -- run directly
echo    dist\%APP_NAME%.zip  -- share this
echo.
echo  Unzip anywhere, double-click %APP_NAME%.exe.
echo  config\ and logs\ folders are created next to the exe
echo  automatically on first run.
echo ============================================================
pause
