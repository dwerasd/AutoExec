@echo off
setlocal EnableExtensions

cd /d "%~dp0"

set "PYTHON_EXE=%~dp0.venv\Scripts\python.exe"
if not exist "%PYTHON_EXE%" set "PYTHON_EXE=python"

echo [1/4] Checking Python...
"%PYTHON_EXE%" --version
if errorlevel 1 goto :error

echo [2/4] Installing build dependencies...
"%PYTHON_EXE%" -m pip install -r requirements.txt PyInstaller
if errorlevel 1 goto :error

echo [3/4] Building autoexec.exe...
"%PYTHON_EXE%" -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name autoexec ^
    --distpath "%~dp0dist" ^
    --workpath "%~dp0build" ^
    --specpath "%~dp0build" ^
    "%~dp0AutoExec.pyw"
if errorlevel 1 goto :error

echo [4/4] Copying runtime files...
for %%F in (
    AutoExec.json
    commands_config.json
    folder_config.json
    registry_config.json
    gitclone.py
    gitsync.py
) do (
    if exist "%~dp0%%F" copy /Y "%~dp0%%F" "%~dp0dist\%%F" >nul
)

if exist "%~dp0AutoExec.db" if not exist "%~dp0dist\AutoExec.db" (
    copy /Y "%~dp0AutoExec.db" "%~dp0dist\AutoExec.db" >nul
)

echo.
echo Build completed: "%~dp0dist\autoexec.exe"
exit /b 0

:error
echo.
echo Build failed. See the error messages above.
exit /b 1
