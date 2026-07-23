@echo off
setlocal EnableExtensions

cd /d "%~dp0"

set "PYTHON_EXE=%~dp0.venv\Scripts\python.exe"
if not exist "%PYTHON_EXE%" set "PYTHON_EXE=python"
set "AUTOEXEC_EXE=%~dp0dist\autoexec.exe"

echo [1/6] Checking Python...
"%PYTHON_EXE%" --version
if errorlevel 1 goto :error

echo [2/6] Stopping running AutoExec...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "try { " ^
    "  $target = [IO.Path]::GetFullPath($env:AUTOEXEC_EXE); " ^
    "  $processes = Get-Process -Name autoexec -ErrorAction SilentlyContinue | Where-Object { $_.Path -and [IO.Path]::GetFullPath($_.Path) -ieq $target }; " ^
    "  if ($processes) { $processes | Stop-Process -Force -ErrorAction Stop; $processes | Wait-Process -ErrorAction SilentlyContinue }; " ^
    "  exit 0 " ^
    "} catch { Write-Error $_; exit 1 }"
if errorlevel 1 goto :error

echo [3/6] Installing build dependencies...
"%PYTHON_EXE%" -m pip install -r requirements.txt PyInstaller
if errorlevel 1 goto :error

echo [4/6] Building autoexec.exe...
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

echo [5/6] Copying runtime files...
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

echo [6/6] Starting AutoExec...
start "" "%AUTOEXEC_EXE%"
if errorlevel 1 goto :error

echo.
echo Build completed and AutoExec started: "%AUTOEXEC_EXE%"
exit /b 0

:error
echo.
echo Build failed. See the error messages above.
exit /b 1
