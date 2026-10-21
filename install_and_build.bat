@echo off
setlocal
cd /d "%~dp0"
set "VENV_PY=%~dp0.venv\Scripts\python.exe"
set "LOG_FILE=%~dp0install_and_build.log"
set "ISCC_EXE="
set "FAIL_STEP=initialization"
>>"%LOG_FILE%" echo.
>>"%LOG_FILE%" echo === Main-only Windows installer build started %date% %time% ===
echo Log file: "%LOG_FILE%"

echo [1/7] Checking Python 3 launcher...
set "FAIL_STEP=Python launcher check"
where py >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
py -3 --version >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

echo [2/7] Checking Inno Setup compiler...
set "FAIL_STEP=Inno Setup compiler check"
if exist "%ProgramFiles%\Inno Setup 7\ISCC.exe" set "ISCC_EXE=%ProgramFiles%\Inno Setup 7\ISCC.exe"
if not defined ISCC_EXE if exist "%ProgramFiles(x86)%\Inno Setup 7\ISCC.exe" set "ISCC_EXE=%ProgramFiles(x86)%\Inno Setup 7\ISCC.exe"
if not defined ISCC_EXE if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC_EXE=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC_EXE if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC_EXE=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not defined ISCC_EXE for /f "delims=" %%I in ('where ISCC.exe 2^>nul') do if not defined ISCC_EXE set "ISCC_EXE=%%I"
if not defined ISCC_EXE (
    echo Inno Setup compiler ISCC.exe was not found.
    echo Install Inno Setup 6 or 7, then run this script again to create the Windows installer.
    >>"%LOG_FILE%" echo ISCC.exe not found. Inno Setup 6 or 7 is required to produce the setup EXE.
    goto failed
)
echo Using Inno Setup compiler: "%ISCC_EXE%"
>>"%LOG_FILE%" echo Using ISCC: "%ISCC_EXE%"

echo [3/7] Creating or reusing the project virtual environment...
set "FAIL_STEP=virtual environment creation"
if not exist "%VENV_PY%" py -3 -m venv "%~dp0.venv" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
if not exist "%VENV_PY%" goto failed

echo [4/7] Installing requirements into .venv...
set "FAIL_STEP=virtual environment pip upgrade"
"%VENV_PY%" -m pip install --upgrade pip >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
set "FAIL_STEP=requirements installation"
"%VENV_PY%" -m pip install -r "%~dp0requirements.txt" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

echo [5/7] Verifying interpreter and PyWin32 COM imports...
set "FAIL_STEP=PyWin32 COM import/version check"
"%VENV_PY%" -c "import sys, importlib.metadata as metadata, pythoncom, win32com.client; print('Python executable:', sys.executable); print('pywin32 version:', metadata.version('pywin32')); print('pythoncom:', pythoncom.__file__); print('win32com.client:', win32com.client.__file__)" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

echo [6/7] Running syntax, self, and static unit checks...
set "FAIL_STEP=Python syntax check"
"%VENV_PY%" -m py_compile main.py wps_automation_workflow.py licensing.py make_fake_data.py wps_source_selection.py test_static.py test_source_selection.py test_selected_processing.py >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
set "FAIL_STEP=application self-check"
"%VENV_PY%" main.py --self-check >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
set "FAIL_STEP=static unit tests"
"%VENV_PY%" -m unittest -v test_static.py test_source_selection.py test_selected_processing.py >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

echo [7/7] Building the Windows application and installer...
set "FAIL_STEP=PyInstaller main application build"
"%VENV_PY%" -m PyInstaller --clean --noconfirm --windowed --onedir --name WpsSheetAutomation --icon="assets\linote-logo.ico" --add-data "assets\linote-logo.png;assets" --add-data "assets\linote-icon.png;assets" --hidden-import=pythoncom --hidden-import=pywintypes --hidden-import=win32timezone --hidden-import=win32crypt --hidden-import=win32com --hidden-import=win32com.client --collect-all win32com --collect-all cryptography main.py >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
if not exist "%~dp0dist\WpsSheetAutomation\WpsSheetAutomation.exe" goto failed
set "FAIL_STEP=Inno Setup installer compilation"
"%ISCC_EXE%" "%~dp0WpsSheetAutomation.iss" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
set "FAIL_STEP=installer output validation"
if not exist "%~dp0dist\WpsSheetAutomation_Setup_v26.exe" goto failed

echo Main application: %~dp0dist\WpsSheetAutomation\WpsSheetAutomation.exe
echo Windows installer: %~dp0dist\WpsSheetAutomation_Setup_v26.exe
echo Main-only installer does not contain the license generator or signing private key.
echo The installer is unsigned; Windows SmartScreen may show a publisher warning.
>>"%LOG_FILE%" echo Installer build completed %date% %time%
pause
endlocal & exit /b 0

:failed
echo.
echo Windows installer build failed at: %FAIL_STEP%
echo Diagnostic log retained at: "%LOG_FILE%"
>>"%LOG_FILE%" echo FAILED at: %FAIL_STEP% (%date% %time%)
echo Check the diagnostic log. Do not launch main.py with system Python.
pause
endlocal & exit /b 1
