@echo off
setlocal EnableExtensions
REM ---------------------------------------------------------------------------
REM  Draftsman - Windows build. Double-click this file.
REM  Installs Python if needed, installs every dependency, and produces
REM  "dist\Draftsman.exe" as a single standalone file.
REM ---------------------------------------------------------------------------
cd /d "%~dp0"
set "LOG=%~dp0build-win-log.txt"
set "VENV=%~dp0.venv-build"
set "EXE=%~dp0dist\Draftsman.exe"
set "REPORT=%~dp0dist\draftsman-selftest.txt"
echo Draftsman build started %DATE% %TIME% > "%LOG%"

echo.
echo ===============================================
echo   Building Draftsman for Windows
echo   Full log: build-win-log.txt
echo ===============================================
echo.

echo [1/6] Finding or installing Python...
REM  ensure_python.ps1 is shared between apps and reads its override from
REM  MAILEX_BUILD_PYTHON; pass this app's own variable through under that name.
if defined DRAFTSMAN_BUILD_PYTHON set "MAILEX_BUILD_PYTHON=%DRAFTSMAN_BUILD_PYTHON%"
for /f "usebackq delims=" %%P in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ensure_python.ps1" 2^>^> "%LOG%"`) do set "PY=%%P"
if not defined PY goto :fail
if not exist "%PY%" goto :fail
echo       using "%PY%"
echo Using Python: "%PY%" >> "%LOG%"

echo [2/6] Creating a clean build environment...
if exist "%VENV%" rmdir /s /q "%VENV%"
"%PY%" -m venv "%VENV%" >> "%LOG%" 2>&1
if errorlevel 1 goto :fail
set "VPY=%VENV%\Scripts\python.exe"

echo [3/6] Installing dependencies (wheels only, no compiling)...
"%VPY%" -m pip install --upgrade pip wheel >> "%LOG%" 2>&1
"%VPY%" -m pip install --only-binary :all: -r "%~dp0requirements.txt" >> "%LOG%" 2>&1
if errorlevel 1 goto :fail
"%VPY%" -m pip install --only-binary :all: "pyinstaller>=6.11,<7" >> "%LOG%" 2>&1
if errorlevel 1 goto :fail

echo [4/6] Checking the code before packaging...
REM  The working directory is already this folder, so '.' is the source tree.
REM  %~dp0 must NOT go inside the Python string: it always ends in a backslash,
REM  which would escape the closing quote and make it an unterminated literal.
"%VPY%" -c "import sys, pkgutil, importlib; sys.path.insert(0,'.'); import draftsmanlib, docx, openpyxl; n=[importlib.import_module(m.name) for m in pkgutil.walk_packages(draftsmanlib.__path__,'draftsmanlib.')]; print('imports ok:', len(n), 'modules')" >> "%LOG%" 2>&1
if errorlevel 1 goto :fail

echo [5/6] Packaging (this takes a minute or two)...
if exist "%~dp0build" rmdir /s /q "%~dp0build"
if exist "%~dp0dist" rmdir /s /q "%~dp0dist"
"%VPY%" -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "Draftsman" ^
  --collect-submodules draftsmanlib ^
  --collect-all docx ^
  --hidden-import docx ^
  --hidden-import openpyxl ^
  --hidden-import tkinter ^
  --hidden-import tkinter.ttk ^
  --hidden-import tkinter.filedialog ^
  --hidden-import tkinter.messagebox ^
  --hidden-import sqlite3 ^
  --exclude-module pytest ^
  "%~dp0draftsman_app.py" >> "%LOG%" 2>&1
if errorlevel 1 goto :fail
if not exist "%EXE%" goto :fail

echo [6/6] Running the self-test on the built program...
echo       It makes the starter library in a temporary folder and fills every template.
echo       A Draftsman window will open for a few seconds - leave it alone
echo       until it closes by itself.
if exist "%REPORT%" del /q "%REPORT%"
pushd "%~dp0dist"
REM  A windowed build has no console, so a plain invocation returns at once and
REM  the report would be read before it exists. start /wait blocks until the
REM  self-test has actually finished.
start "Draftsman self-test" /wait "%EXE%" selftest
popd
if not exist "%REPORT%" goto :noreport
echo.
type "%REPORT%"
findstr /C:"SELF-TEST PASSED" "%REPORT%" >nul
if errorlevel 1 goto :selftestfail

echo.
echo ===============================================
echo   DONE.  dist\Draftsman.exe is ready.
echo ===============================================
echo.
pause
exit /b 0

:noreport
echo.
echo ===============================================
echo   BUILT, BUT THE PROGRAM DID NOT RUN ITS SELF-TEST.
echo   It probably does not start. Look for
echo   dist\draftsman-crash.log
echo ===============================================
echo.
pause
exit /b 2

:selftestfail
echo.
echo ===============================================
echo   BUILT, BUT THE SELF-TEST REPORTED PROBLEMS.
echo   See dist\draftsman-selftest.txt above.
echo ===============================================
echo.
pause
exit /b 2

:fail
echo.
echo ===============================================
echo   BUILD FAILED.  Last 40 lines of the log:
echo ===============================================
powershell -NoProfile -Command "Get-Content -Tail 40 -LiteralPath '%LOG%'"
echo.
echo Full log: "%LOG%"
echo.
pause
exit /b 1
