@echo off
setlocal EnableExtensions

cd /d "%~dp0"

set "PYTHON_EXE=%~dp0.venv\Scripts\python.exe"
if not exist "%PYTHON_EXE%" set "PYTHON_EXE=python"
set "AUTOEXEC_DIR=%~dp0"
set "AUTOEXEC_EXE=%~dp0AutoExec.exe"

echo [1/6] Checking Python...
"%PYTHON_EXE%" --version
if errorlevel 1 goto :error

echo [2/6] Stopping running AutoExec...
rem exit 3 = process was running and killed (used to decide restart after build)
set "WAS_RUNNING=0"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "try { " ^
    "  $root = [IO.Path]::GetFullPath($env:AUTOEXEC_DIR); " ^
    "  $processes = Get-Process -Name AutoExec -ErrorAction SilentlyContinue | Where-Object { $_.Path -and [IO.Path]::GetFullPath($_.Path).StartsWith($root, [StringComparison]::OrdinalIgnoreCase) }; " ^
    "  if ($processes) { $processes | Stop-Process -Force -ErrorAction Stop; $processes | Wait-Process -ErrorAction SilentlyContinue; exit 3 }; " ^
    "  exit 0 " ^
    "} catch { Write-Error $_; exit 1 }"
if "%errorlevel%"=="1" goto :error
if "%errorlevel%"=="3" set "WAS_RUNNING=1"

echo [3/6] Installing build dependencies...
"%PYTHON_EXE%" -m pip install -r requirements.txt PyInstaller
if errorlevel 1 goto :error

echo [4/6] Building AutoExec.exe...
"%PYTHON_EXE%" -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name AutoExec ^
    --distpath "%~dp0." ^
    --workpath "%~dp0build" ^
    --specpath "%~dp0build" ^
    "%~dp0AutoExec.pyw"
if errorlevel 1 goto :error

echo [5/6] Removing stale dist build...
if exist "%~dp0dist\autoexec.exe" del /f /q "%~dp0dist\autoexec.exe"

if "%WAS_RUNNING%"=="1" (
    echo [6/6] Restarting AutoExec...
    start "" "%AUTOEXEC_EXE%"
) else (
    echo [6/6] AutoExec was not running - skipping restart.
)

echo.
echo Build completed: "%AUTOEXEC_EXE%"
exit /b 0

:error
echo.
echo Build failed. See the error messages above.
exit /b 1
