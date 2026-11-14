@echo off
setlocal
cd /d "%~dp0"
set "INVENTARIS_PYTHON="
for /d %%D in ("C:\Program Files\WindowsApps\OpenAI.CodexPrimaryRuntime.*") do if exist "%%~fD\dependencies\python\python.exe" set "INVENTARIS_PYTHON=%%~fD\dependencies\python\python.exe"
rem Kies bij voorkeur een Python-installatie waarin pyarrow beschikbaar is.
py -3 -c "import pyarrow" >nul 2>nul
if not errorlevel 1 goto launcher
python -c "import pyarrow" >nul 2>nul
if not errorlevel 1 goto python
if defined INVENTARIS_PYTHON goto bundled
where py >nul 2>nul
if not errorlevel 1 goto launcher
echo Python is niet gevonden. Open Codex om de inventarispagina bij te werken.
pause
exit /b 1
:bundled
"%INVENTARIS_PYTHON%" "%~dp0bijwerken.py"
goto finished
:launcher
py -3 "%~dp0bijwerken.py"
goto finished
:python
python "%~dp0bijwerken.py"
:finished
if errorlevel 1 goto failed
echo Vernieuw index.html in je browser om de nieuwe gegevens te zien.
pause
exit /b 0
:failed
echo De vorige inventarispagina is behouden.
pause
exit /b 1
