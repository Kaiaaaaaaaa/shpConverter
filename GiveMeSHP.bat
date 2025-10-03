@echo off
setlocal
pushd "%~dp0"
set "PY=%~dp0venv\Scripts\python.exe"
if not exist "%PY%" exit /b 1

"%PY%" "%~dp0dxf2shp.py"
if errorlevel 1 exit /b %errorlevel%

"%PY%" "%~dp0makePrj4shp.py"
if errorlevel 1 exit /b %errorlevel%

pause
