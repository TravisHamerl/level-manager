@echo off
:: Launch Level Manager — works on any machine with Python 3 installed
:: Tries common Python locations, py launcher, then PATH

set "SCRIPT=%~dp0level_manager.py"

:: Try common install locations
for %%P in (
    "C:\Program Files\Python314\pythonw.exe"
    "C:\Program Files\Python313\pythonw.exe"
    "C:\Program Files\Python312\pythonw.exe"
    "C:\Program Files\Python311\pythonw.exe"
    "C:\Program Files\Python310\pythonw.exe"
    "C:\Python314\pythonw.exe"
    "C:\Python313\pythonw.exe"
    "C:\Python312\pythonw.exe"
    "%LocalAppData%\Programs\Python\Python314\pythonw.exe"
    "%LocalAppData%\Programs\Python\Python313\pythonw.exe"
    "%LocalAppData%\Programs\Python\Python312\pythonw.exe"
    "%LocalAppData%\Programs\Python\Python311\pythonw.exe"
    "%LocalAppData%\Programs\Python\Python310\pythonw.exe"
) do (
    if exist %%P (
        start "" %%P "%SCRIPT%"
        exit /b 0
    )
)

:: Try py launcher (installed with Python even without PATH)
where pyw >nul 2>&1
if %errorlevel%==0 (
    start "" pyw "%SCRIPT%"
    exit /b 0
)

where py >nul 2>&1
if %errorlevel%==0 (
    start "" py "%SCRIPT%"
    exit /b 0
)

:: Fall back to PATH
where pythonw >nul 2>&1
if %errorlevel%==0 (
    start "" pythonw "%SCRIPT%"
    exit /b 0
)

where python >nul 2>&1
if %errorlevel%==0 (
    start "" python "%SCRIPT%"
    exit /b 0
)

echo Python not found! Please install Python 3.10+ from python.org
echo Make sure to check "Add Python to PATH" during installation.
pause
