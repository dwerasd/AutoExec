@echo off
setlocal EnableExtensions

cd /d "%~dp0"

set "PYTHON_EXE=%~dp0.venv\Scripts\python.exe"
if not exist "%PYTHON_EXE%" set "PYTHON_EXE=python"
set "AUTOEXEC_EXE=%~dp0dist\autoexec.exe"

echo [1/5] Checking Python...
"%PYTHON_EXE%" --version
if errorlevel 1 goto :error

echo [2/5] Stopping running AutoExec...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "try { " ^
    "  $target = [IO.Path]::GetFullPath($env:AUTOEXEC_EXE); " ^
    "  $processes = Get-Process -Name autoexec -ErrorAction SilentlyContinue | Where-Object { $_.Path -and [IO.Path]::GetFullPath($_.Path) -ieq $target }; " ^
    "  if ($processes) { $processes | Stop-Process -Force -ErrorAction Stop; $processes | Wait-Process -ErrorAction SilentlyContinue }; " ^
    "  exit 0 " ^
    "} catch { Write-Error $_; exit 1 }"
if errorlevel 1 goto :error

echo [3/5] Installing build dependencies...
"%PYTHON_EXE%" -m pip install -r requirements.txt PyInstaller
if errorlevel 1 goto :error

echo [4/5] Building autoexec.exe...
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

echo [5/5] Copying runtime files...
for %%F in (
    commands_config.json
    folder_config.json
    registry_config.json
    gitclone.py
    gitsync.py
) do (
    if exist "%~dp0%%F" copy /Y "%~dp0%%F" "%~dp0dist\%%F" >nul
)

rem AutoExec.json/AutoExec.db 는 exe 런타임 상태 파일 — 최초 1회만 시드 (덮어쓰면 창 위치 등 롤백)
for %%F in (
    AutoExec.json
    AutoExec.db
) do (
    if exist "%~dp0%%F" if not exist "%~dp0dist\%%F" (
        copy /Y "%~dp0%%F" "%~dp0dist\%%F" >nul
    )
)

echo.
echo Build completed: "%AUTOEXEC_EXE%"
exit /b 0

:error
echo.
echo Build failed. See the error messages above.
exit /b 1
