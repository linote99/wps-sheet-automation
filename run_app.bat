@echo off
setlocal
cd /d "%~dp0"
set "VENV_PY=%~dp0.venv\Scripts\python.exe"

if not exist "%VENV_PY%" (
  echo The project virtual environment was not found:
  echo "%VENV_PY%"
  echo First run install_and_build.bat from this folder.
  echo If installation failed, inspect build_main.log.
  pause
  endlocal & exit /b 1
)

echo Checking PyWin32 imports from the project virtual environment...
"%VENV_PY%" -c "import sys, pythoncom, win32com.client; print('Python executable:', sys.executable); print('pythoncom:', pythoncom.__file__)"
if errorlevel 1 (
  echo.
  echo PyWin32 COM imports failed in this project's .venv.
  echo Run install_and_build.bat and review build_main.log if it fails.
  pause
  endlocal & exit /b 1
)

"%VENV_PY%" "%~dp0main.py" %*
set "APP_EXIT_CODE=%ERRORLEVEL%"
endlocal & exit /b %APP_EXIT_CODE%
