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
rem --onedir 고정: --onefile 은 실행할 때마다 %%TEMP%%\_MEIxxxxx 로 압축을 풀고
rem 종료 시 지운다. 그 폴더의 DLL 을 아직 물고 있는 인스턴스가 있으면 삭제가
rem 실패하고 부트로더가 "Failed to remove temporary directory" 경고 창을 띄우는데,
rem 무인 RPA 에서는 사람이 확인을 누를 때까지 멈춘다. onedir 은 임시 폴더를 쓰지 않는다.
"%PYTHON_EXE%" -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onedir ^
    --windowed ^
    --name AutoExec ^
    --distpath "%~dp0build\dist" ^
    --workpath "%~dp0build" ^
    --specpath "%~dp0build" ^
    "%~dp0AutoExec.pyw"
if errorlevel 1 goto :error

echo [5/6] Deploying build output...
rem AutoExec 은 SCRIPT_DIR(= AutoExec.exe 가 있는 폴더)에서 DB/.env/JSON 을 찾는다.
rem 따라서 exe 는 프로젝트 루트에 두고 런타임 폴더(_internal)만 그 옆에 미러링한다.
copy /y "%~dp0build\dist\AutoExec\AutoExec.exe" "%AUTOEXEC_EXE%" >nul
if errorlevel 1 goto :error
robocopy "%~dp0build\dist\AutoExec\_internal" "%~dp0_internal" /MIR /NFL /NDL /NJH /NJS /NP >nul
if errorlevel 8 goto :error
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
