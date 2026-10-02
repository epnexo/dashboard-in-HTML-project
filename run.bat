@echo off
rem Opens the report window. Double-click to use.
cd /d "%~dp0"
where pyw >nul 2>nul && (start "" pyw -3 scripts\app.py & exit /b)
where pythonw >nul 2>nul && (start "" pythonw scripts\app.py & exit /b)
if exist "%LOCALAPPDATA%\Programs\Python\Python312\pythonw.exe" (start "" "%LOCALAPPDATA%\Programs\Python\Python312\pythonw.exe" scripts\app.py & exit /b)
echo Python was not found. Install it from https://www.python.org/downloads/ and try again.
pause
