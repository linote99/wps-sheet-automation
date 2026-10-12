@echo off
setlocal
cd /d "%~dp0"
set "VENV_PY=%~dp0.venv\Scripts\python.exe"
set "LOG_FILE=%~dp0build_main.log"
set "FAIL_STEP=initialization"
>>"%LOG_FILE%" echo.
>>"%LOG_FILE%" echo === Main app build started %date% %time% ===
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

set "FAIL_STEP=requirements installation"
"%VENV_PY%" -m pip install --upgrade pip >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
"%VENV_PY%" -m pip install -r "%~dp0requirements.txt" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

set "FAIL_STEP=PyWin32 COM import check"
"%VENV_PY%" -c "import sys, importlib.metadata as metadata, pythoncom, win32com.client; import win32crypt; print('Python executable:', sys.executable); print('pywin32 version:', metadata.version('pywin32')); print('PyWin32 COM imports OK')" >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

set "FAIL_STEP=syntax check"
"%VENV_PY%" -m py_compile main.py wps_automation_workflow.py licensing.py test_static.py >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
set "FAIL_STEP=application self-check"
"%VENV_PY%" main.py --self-check >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed
set "FAIL_STEP=static regression tests"
"%VENV_PY%" -m unittest -v test_static.py >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

set "FAIL_STEP=PyInstaller main application build"
"%VENV_PY%" -m PyInstaller --clean --noconfirm --windowed --onedir --name WpsSheetAutomation --hidden-import=pythoncom --hidden-import=pywintypes --hidden-import=win32timezone --hidden-import=win32crypt --hidden-import=win32com --hidden-import=win32com.client --collect-all win32com --collect-all cryptography main.py >>"%LOG_FILE%" 2>&1
if errorlevel 1 goto failed

set "ISCC_EXE="
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC_EXE=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC_EXE=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not defined ISCC_EXE (
  for /f "delims=" %%I in ('where ISCC.exe 2^>nul') do (
    if not defined ISCC_EXE set "ISCC_EXE=%%I"
  )
)
if defined ISCC_EXE (
  set "FAIL_STEP=Inno Setup installer compilation"
  "%ISCC_EXE%" /Qp "%~dp0WpsSheetAutomation.iss" >>"%LOG_FILE%" 2>&1
  if errorlevel 1 goto failed
  echo Installer built: "%~dp0dist\WpsSheetAutomation_Setup_v13.exe"
  >>"%LOG_FILE%" echo Installer built: "%~dp0dist\WpsSheetAutomation_Setup_v13.exe"
) else (
  echo Inno Setup 6 was not found. The app folder is built, but no Setup EXE was created.
  echo Install Inno Setup 6 and rerun this script to create the installer.
  >>"%LOG_FILE%" echo Installer not built: Inno Setup 6 compiler ISCC.exe was not found.
)

echo Main application built: "%~dp0dist\WpsSheetAutomation\WpsSheetAutomation.exe"
echo This package contains no license-generator code or issuer private key.
>>"%LOG_FILE%" echo Build completed %date% %time%
pause
endlocal & exit /b 0

:failed
echo.
echo Build failed at: %FAIL_STEP%
echo Review log: "%LOG_FILE%"
>>"%LOG_FILE%" echo FAILED at: %FAIL_STEP% (%date% %time%)
pause
endlocal & exit /b 1
