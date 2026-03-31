@echo off
echo ============================================
echo  Contract Generator -- Build EXE
echo ============================================

echo.
echo [1/3] Installing dependencies...
pip install python-docx openpyxl pyinstaller

echo.
echo [2/3] Building executable...
pyinstaller ^
  --onefile ^
  --windowed ^
  --name "ContractGenerator" ^
  --icon NONE ^
  app.py

echo.
echo [3/3] Done!
echo.
echo Your EXE is at:  dist\ContractGenerator.exe
echo.
echo NOTE: LibreOffice must be installed on the target machine for PDF export.
echo       Download from: https://www.libreoffice.org/download/download-libreoffice/
echo.
pause
