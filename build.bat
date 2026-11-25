@echo off
REM Builds dist\CMTS_Command_Center.exe in a clean virtual environment.
REM Run from the repo folder:  build.bat
setlocal
cd /d "%~dp0"

where py >nul 2>nul && (set PY=py -3) || (set PY=python)

if not exist .venv (
    echo [1/4] Creating virtual environment...
    %PY% -m venv .venv || goto :fail
)

echo [2/4] Installing dependencies...
call .venv\Scripts\activate.bat || goto :fail
python -m pip install --upgrade pip >nul
python -m pip install --upgrade pyinstaller customtkinter paramiko matplotlib openpyxl || goto :fail

echo [3/4] Checking the app imports...
python -c "import customtkinter, paramiko, matplotlib, openpyxl; import py_compile; py_compile.compile('CMTS_Command_Center.py', doraise=True)" || goto :fail

echo [4/4] Building exe...
if exist build rmdir /s /q build
if exist dist\CMTS_Command_Center.exe del /f /q dist\CMTS_Command_Center.exe || (
    echo.
    echo dist\CMTS_Command_Center.exe is locked. Close the app and try again.
    goto :fail
)
pyinstaller --clean --noconfirm CMTS_Command_Center.spec || goto :fail

echo.
echo Done: %cd%\dist\CMTS_Command_Center.exe
endlocal
exit /b 0

:fail
echo.
echo BUILD FAILED - copy everything above this line and send it over.
endlocal
exit /b 1
