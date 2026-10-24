@echo off
setlocal
cd /d "%~dp0"
echo.
echo [1/3] Checking Python 3.11 or newer...
py -3 -c "import sys; sys.exit(sys.version_info < (3, 11))" >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=py -3"
    goto :install
)
python -c "import sys; sys.exit(sys.version_info < (3, 11))" >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=python"
    goto :install
)
echo.
echo [ERROR] Python 3.11 or newer is not available.
echo Install Python 3.11 or newer, then run this file again.
pause
exit /b 1

:install
echo [2/3] Installing Python packages...
%PY_CMD% -m pip --version >nul 2>&1
if errorlevel 1 (
    %PY_CMD% -m ensurepip --upgrade
)
%PY_CMD% -m pip install -r "requirements.txt"
if errorlevel 1 (
    echo.
    echo [ERROR] Package installation failed. See the error above.
    pause
    exit /b 1
)
echo.
echo [3/3] Installation finished successfully.
echo Next: double-click the Run batch file in this folder.
pause
exit /b 0
