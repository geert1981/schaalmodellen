@echo off
setlocal
cd /d "%~dp0"
py -3 -c "import pyarrow" >nul 2>nul
if not errorlevel 1 goto launcher
python -c "import pyarrow" >nul 2>nul
if not errorlevel 1 goto python
set "INVENTARIS_PYTHON="
for /d %%D in ("C:\Program Files\WindowsApps\OpenAI.CodexPrimaryRuntime.*") do if exist "%%~fD\dependencies\python\python.exe" set "INVENTARIS_PYTHON=%%~fD\dependencies\python\python.exe"
if defined INVENTARIS_PYTHON goto bundled
where py >nul 2>nul
if not errorlevel 1 goto launcher
where python >nul 2>nul
if not errorlevel 1 goto python
echo Python is niet gevonden.
pause
exit /b 1
:launcher
py -3 "%~dp0server.py"
goto finished
:python
python "%~dp0server.py"
goto finished
:bundled
"%INVENTARIS_PYTHON%" "%~dp0server.py"
:finished
if not errorlevel 1 exit /b 0
echo De app is gestopt door een fout. Zie de melding hierboven.
pause
exit /b 1
