@echo off
setlocal
cd /d "%~dp0"

echo [1/5] Checking Python 3...
where py >nul 2>nul
if errorlevel 1 (
  echo Python Launcher 'py' was not found. Install Python 3.10 or newer first.
  exit /b 1
)
py -3 --version
if errorlevel 1 exit /b 1

echo [2/5] Creating virtual environment...
if not exist ".venv\Scripts\python.exe" py -3 -m venv .venv
if errorlevel 1 exit /b 1
call ".venv\Scripts\activate.bat"
if errorlevel 1 exit /b 1

echo [3/5] Installing dependencies...
python -m pip install --upgrade pip
if errorlevel 1 exit /b 1
python -m pip install -r requirements.txt
if errorlevel 1 exit /b 1

echo [4/5] Running static checks...
python -m py_compile main.py workflow.py make_fake_data.py test_static.py
if errorlevel 1 exit /b 1
python main.py --self-check
if errorlevel 1 exit /b 1
python -m unittest -v test_static.py
if errorlevel 1 exit /b 1

echo [5/5] Building Windows application folder...
python -m PyInstaller --clean --noconfirm --windowed --onedir --name WpsSheetAutomation --hidden-import=pythoncom --hidden-import=pywintypes --hidden-import=win32timezone --hidden-import=win32com.client main.py
if errorlevel 1 exit /b 1

echo Build complete: %~dp0dist\WpsSheetAutomation\WpsSheetAutomation.exe
echo WPS Table must still be installed on the target Windows PC. COM operation has not been tested by this build script.
endlocal
