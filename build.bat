@echo off
setlocal
cd /d "%~dp0"

echo ============================================
echo  Contract Generator -- Build EXE
echo ============================================

if not exist "app.py" (
  echo.
  echo ERROR: app.py was not found in this folder.
  goto :fail
)

if not exist "requirements.txt" (
  echo.
  echo ERROR: requirements.txt was not found in this folder.
  goto :fail
)

if not exist "Workflow Automation template.docx" (
  echo.
  echo ERROR: Workflow Automation template.docx was not found in this folder.
  goto :fail
)

where python >nul 2>nul
if errorlevel 1 (
  echo.
  echo ERROR: Python is not installed or not on PATH.
  echo Install Python 3 on this Windows machine, then rerun this file.
  goto :fail
)

echo.
echo [1/5] Creating virtual environment...
if not exist ".venv\Scripts\python.exe" (
  python -m venv .venv
  if errorlevel 1 goto :fail
)

echo.
echo [2/5] Installing dependencies...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :fail
".venv\Scripts\python.exe" -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto :fail

echo.
echo [3/5] Building application folder...
set "PYI_WORK=%TEMP%\ContractGenerator-pyinstaller\work"
set "PYI_SPEC=%TEMP%\ContractGenerator-pyinstaller\spec"
set "TEMPLATE_PATH=%CD%\Workflow Automation template.docx"
if exist "%PYI_WORK%" rmdir /s /q "%PYI_WORK%"
if exist "%PYI_SPEC%" rmdir /s /q "%PYI_SPEC%"

".venv\Scripts\python.exe" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --distpath "dist" ^
  --workpath "%PYI_WORK%" ^
  --specpath "%PYI_SPEC%" ^
  --onedir ^
  --windowed ^
  --name "ContractGenerator" ^
  --add-data "%TEMPLATE_PATH%;." ^
  app.py
if errorlevel 1 goto :fail

if not exist "dist\ContractGenerator\ContractGenerator.exe" (
  echo.
  echo ERROR: Build finished but dist\ContractGenerator\ContractGenerator.exe was not created.
  goto :fail
)

echo.
echo [4/5] Creating shareable zip...
if exist "dist\ContractGenerator-win.zip" del /f /q "dist\ContractGenerator-win.zip"
powershell -NoProfile -NonInteractive -Command "Compress-Archive -Path 'dist\ContractGenerator' -DestinationPath 'dist\ContractGenerator-win.zip' -Force"
if errorlevel 1 goto :fail

echo.
echo [5/5] Done!
echo.
echo Your app folder is at:  dist\ContractGenerator
echo Main EXE is at:         dist\ContractGenerator\ContractGenerator.exe
echo Shareable zip is at:    dist\ContractGenerator-win.zip
echo.
echo The app folder includes Python dependencies and the default template.
echo DOCX generation and export work on a fresh Windows machine with nothing else installed.
echo PDF export tries Microsoft Word first on Windows.
echo If Word is not installed, it falls back to LibreOffice if available.
echo If neither is available, export falls back to DOCX.
echo This one-folder build is also generally less likely to be blocked than --onefile.
echo.
pause
exit /b 0

:fail
echo.
echo Build failed.
echo If this keeps happening, run this file from Command Prompt and share the error output.
echo.
pause
exit /b 1
