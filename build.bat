@echo off
echo ============================================
echo  Contract Generator -- Build EXE
echo ============================================

echo.
echo [1/3] Installing dependencies...
pip install -r requirements.txt pyinstaller

echo.
echo [2/3] Building executable...
pyinstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name "ContractGenerator" ^
  --add-data "Workflow Automation template.docx;." ^
  --icon NONE ^
  app.py

echo.
echo [3/3] Done!
echo.
echo Your EXE is at:  dist\ContractGenerator.exe
echo.
echo The EXE includes Python dependencies and the default template.
echo DOCX generation and export work on a fresh Windows machine with nothing else installed.
echo PDF export tries Microsoft Word first on Windows.
echo If Word is not installed, it falls back to LibreOffice if available.
echo If neither is available, export falls back to DOCX.
echo.
pause
