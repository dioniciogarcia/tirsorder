@echo off
title Building NuORDER Client
cd /d "%~dp0"
echo ============================================
echo   Building "NuORDER Client.exe"  (3-5 minutes)
echo ============================================
echo.
py --version >nul 2>&1
if errorlevel 1 (
  echo Python was not found. Install it from https://www.python.org/downloads/
  echo IMPORTANT: tick "Add python.exe to PATH" on the first installer screen.
  pause
  exit /b 1
)
py -m pip install --upgrade pip
py -m pip install --upgrade openpyxl pillow pymupdf pyinstaller
if errorlevel 1 (
  echo Could not install the required packages. Check your internet connection.
  pause
  exit /b 1
)
if exist "NuORDER Client.exe" del /q "NuORDER Client.exe"
py -m PyInstaller --noconfirm --clean --onefile --windowed --name "NuORDER Client" --hidden-import PIL.Image --hidden-import PIL.ImageOps --hidden-import PIL.JpegImagePlugin --hidden-import PIL.PngImagePlugin --hidden-import pymupdf --hidden-import fitz --collect-submodules pymupdf nuorder_client.py
if errorlevel 1 (
  echo The build failed. Copy the text above and send it to Claude.
  pause
  exit /b 1
)
rem --- keep only the finished program, remove PyInstaller's work folders ---
move /y "dist\NuORDER Client.exe" "NuORDER Client.exe" >nul
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
rmdir /s /q __pycache__ 2>nul
del /q "NuORDER Client.spec" 2>nul
echo.
echo ============================================
echo   DONE!  Your program is here:
echo   %cd%\NuORDER Client.exe
echo   It is one single file - copy it anywhere you like.
echo ============================================
pause
