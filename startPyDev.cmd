@echo off
REM Get current script directory
set "PROJECT_DIR=%~dp0"
set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"  REM Remove trailing backslash

REM Extract drive root (e.g. D:\)
for %%I in ("%PROJECT_DIR%") do set "DRIVE_ROOT=%%~dI\"

REM Check if project is directly in root
if /I "%PROJECT_DIR%"=="%DRIVE_ROOT%" (
    echo Project is located directly in the root directory.
    set "VENV_DIR=%DRIVE_ROOT%venv"
) else (
    echo Project is located in a subdirectory.
    set "VENV_DIR=%PROJECT_DIR%_venv"
)

REM Check if virtual environment exists
if not exist "%VENV_DIR%\Scripts\activate" (
    echo Creating virtual environment at: %VENV_DIR%
    python -m venv "%VENV_DIR%"
)

REM Activate the virtual environment
call "%VENV_DIR%\Scripts\activate"

if exist requirements.txt (
	pip install -r requirements.txt
)

echo "deactivate exits the venv, exit closes the cmd"
cmd /k