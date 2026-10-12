@echo off
setlocal
cd /d "%~dp0"
set "VENV_PY=%~dp0.venv\Scripts\python.exe"
set "LOG_FILE=%~dp0install_and_build.log"
set "FAIL_STEP=initialization"
>>"%LOG_FILE%" echo.
>>"%LOG_FILE%" echo === Main-only build started %date% %time% ===
echo Log file: "%LOG_FILE%"

set "FAIL_STEP=Python launcher check"
where py >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
py -3 --version >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

set "FAIL_STEP=virtual environment creation"
if not exist "%VENV_PY%" py -3 -m venv "%~dp0.venv" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
if not exist "%VENV_PY%" goto failed

set "FAIL_STEP=pip upgrade"
"%VENV_PY%" -m pip install --upgrade pip >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
set "FAIL_STEP=requirements installation"
"%VENV_PY%" -m pip install -r "%~dp0requirements.txt" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

set "FAIL_STEP=PyWin32 COM import/version check"
"%VENV_PY%" -c "import sys, importlib.metadata as metadata, pythoncom, win32com.client; print('Python executable:', sys.executable); print('pywin32 version:', metadata.version('pywin32')); print('pythoncom:', pythoncom.__file__); print('win32com.client:', win32com.client.__file__)" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

set "FAIL_STEP=Python syntax check"
"%VENV_PY%" -m py_compile main.py wps_automation_workflow.py licensing.py make_fake_data.py test_static.py >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
set "FAIL_STEP=application self-check"
"%VENV_PY%" main.py --self-check >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
set "FAIL_STEP=static tests"
"%VENV_PY%" -m unittest -v test_static.py >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

set "FAIL_STEP=PyInstaller main application build"
"%VENV_PY%" -m PyInstaller --clean --noconfirm --windowed --onedir --name WpsSheetAutomation --icon="assets\linote-logo.ico" --add-data "assets\linote-logo.png;assets" --add-data "assets\linote-icon.png;assets" --hidden-import=pythoncom --hidden-import=pywintypes --hidden-import=win32timezone --hidden-import=win32crypt --hidden-import=win32com --hidden-import=win32com.client --collect-all win32com --collect-all cryptography main.py >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

echo Build complete: %~dp0dist\WpsSheetAutomation\WpsSheetAutomation.exe
echo This script builds only the main application; it does not build a key generator or copy any signing private key.
echo Real WPS COM operation and the packaged EXE still require Windows acceptance testing.
>>"%LOG_FILE%" echo Build completed %date% %time%
pause
endlocal & exit /b 0

:failed
echo.
echo Build failed at: %FAIL_STEP%
echo Diagnostic log retained at: "%LOG_FILE%"
>>"%LOG_FILE%" echo FAILED at: %FAIL_STEP% (%date% %time%)
echo Review the log, correct the reported issue, and retry.
pause
endlocal & exit /b 1
