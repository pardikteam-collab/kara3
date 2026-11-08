@echo off
setlocal
cd /d "%~dp0"
if not exist "core\__init__.py" (
  echo ERROR: Required core folder is missing. Extract the complete source ZIP.
  pause
  exit /b 1
)
py -m pip install --upgrade pip
py -m pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --clean --windowed --onedir --name KaraShiftManager --paths . --hidden-import core --hidden-import core.excel_reader --hidden-import core.mapping --hidden-import core.automation --collect-all PySide6 --collect-all shiboken6 main.py
if errorlevel 1 (
  echo Build failed.
  pause
  exit /b 1
)
echo.
echo Build completed.
echo Run: dist\KaraShiftManager\KaraShiftManager.exe
echo IMPORTANT: Keep the entire KaraShiftManager folder together.
pause
