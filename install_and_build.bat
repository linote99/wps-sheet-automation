@echo off
setlocal
cd /d "%~dp0"
set "VENV_PY=%~dp0.venv\Scripts\python.exe"
set "LOG_FILE=%~dp0install_and_build.log"
set "FAIL_STEP=initialization"
>>"%LOG_FILE%" echo.
>>"%LOG_FILE%" echo === Installer started %date% %time% ===
echo Log file: "%LOG_FILE%"

echo [1/6] Checking Python 3 launcher...
set "FAIL_STEP=Python launcher check"
where py >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
py -3 --version >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

echo [2/6] Creating or reusing the project virtual environment...
set "FAIL_STEP=virtual environment creation"
if not exist "%VENV_PY%" py -3 -m venv "%~dp0.venv" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
if not exist "%VENV_PY%" goto failed

echo [3/6] Installing requirements into .venv...
set "FAIL_STEP=virtual environment pip upgrade"
"%VENV_PY%" -m pip install --upgrade pip >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
set "FAIL_STEP=requirements installation"
"%VENV_PY%" -m pip install -r "%~dp0requirements.txt" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

echo [4/6] Verifying interpreter and PyWin32 COM imports...
set "FAIL_STEP=PyWin32 COM import/version check"
"%VENV_PY%" -c "import sys, importlib.metadata as metadata, pythoncom, win32com.client; print('Python executable:', sys.executable); print('pywin32 version:', metadata.version('pywin32')); print('pythoncom:', pythoncom.__file__); print('win32com.client:', win32com.client.__file__)" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

echo [5/6] Running syntax, self, and unit checks...
set "FAIL_STEP=Python syntax check"
"%VENV_PY%" -m py_compile main.py wps_automation_workflow.py licensing.py keygen.py make_fake_data.py test_static.py test_licensing.py >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
set "FAIL_STEP=application self-check"
"%VENV_PY%" main.py --self-check >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
set "FAIL_STEP=unit tests"
"%VENV_PY%" -m unittest -v test_static.py test_licensing.py >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

echo [6/6] Building Windows application and license-generator folders...
FAIL_STEP=PyInstaller build
"%VENV_PY%" -m PyInstaller --clean --noconfirm --windowed --onedir --name WpsSheetAutomation --hidden-import=pythoncom --hidden-import=pywintypes --hidden-import=win32timezone --hidden-import=win32crypt --hidden-import=win32com --hidden-import=win32com.client --collect-all win32com --collect-all cryptography main.py >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

FAIL_STEP=License generator PyInstaller build
"%VENV_PY%" -m PyInstaller --noconfirm --windowed --onedir --name WpsSheetAutomation_KeyGen --collect-all cryptography keygen.py >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
if not exist "%~dp0dist\WpsSheetAutomation_KeyGen\licensing_private" mkdir "%~dp0dist\WpsSheetAutomation_KeyGen\licensing_private" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
copy /Y "%~dp0licensing_private\issuer_private_key.pem" "%~dp0dist\WpsSheetAutomation_KeyGen\licensing_private\issuer_private_key.pem" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
copy /Y "%~dp0ufitools_keygen_plugin.html" "%~dp0dist\WpsSheetAutomation_KeyGen\ufitools_keygen_plugin.html" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

echo Build complete: %~dp0dist\WpsSheetAutomation\WpsSheetAutomation.exe
echo Key generator: %~dp0dist\WpsSheetAutomation_KeyGen\WpsSheetAutomation_KeyGen.exe
echo UFI-TOOLS plugin: %~dp0ufitools_keygen_plugin.html
echo IMPORTANT: Keep the KeyGen folder and issuer_private_key.pem private; do not send them to licensees.
echo Source launch entry: %~dp0run_app.bat
echo WPS Table must be installed on the target Windows PC. Real WPS COM operation and the packaged EXE still require Windows acceptance testing.
>>"%LOG_FILE%" echo Build completed %date% %time%
pause
endlocal & exit /b 0

:failed
echo.
echo Installer/build failed at: %FAIL_STEP%
echo Diagnostic log retained at: "%LOG_FILE%"
>>"%LOG_FILE%" echo FAILED at: %FAIL_STEP% (%date% %time%)
echo Do not launch main.py with system Python. Review the log, then retry this installer.
pause
endlocal & exit /b 1
