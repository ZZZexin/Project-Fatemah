@echo off
echo ============================================================
echo  TV Pipeline -- build TvPipeline.exe
echo ============================================================
echo.

:: Check Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: python not found on PATH.
    pause & exit /b 1
)

:: Install / upgrade PyInstaller quietly
echo Installing PyInstaller...
pip install pyinstaller --quiet --upgrade

echo.
echo Building (this takes 1-3 minutes the first time)...
pyinstaller build.spec --noconfirm --clean

if errorlevel 1 (
    echo.
    echo BUILD FAILED -- see output above.
    pause & exit /b 1
)

echo.
echo ============================================================
echo  Done!  Output: dist\TvPipeline.exe
echo.
echo  Copy TvPipeline.exe to any department PC.
echo  On first launch it creates a  config\  folder next to it
echo  for saved settings and a  logs\  folder for convert logs.
echo ============================================================
pause
