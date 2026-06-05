@echo off
REM UDF Wizard - Windows baslatici
set "APPDIR=%~dp0"
set "APPDIR=%APPDIR:~0,-1%"
for /f "tokens=*" %%i in ('powershell -NoProfile -Command "[Environment]::GetFolderPath('Desktop')"') do set "DESKTOP=%%i"
set "SHORTCUT=%DESKTOP%\UDF Wizard.lnk"
set "ICO=%APPDIR%\UDF_LOGO.ico"
set "APP_SCRIPT=%APPDIR%\udf_donusturucu.py"

python --version >nul 2>&1
if errorlevel 1 (
    echo HATA: Python bulunamadi.
    echo https://www.python.org/downloads/ adresinden indirip kurun.
    echo Kurulum sirasinda "Add Python to PATH" secenegini isaretleyin.
    pause
    exit /b
)

pip install customtkinter python-docx pdfplumber reportlab pillow --quiet

if not exist "%ICO%" (
    python -c "from PIL import Image; img=Image.open(r'%APPDIR%\UDF_LOGO.png').convert('RGBA'); img.save(r'%ICO%', format='ICO', sizes=[(256,256),(128,128),(64,64),(48,48),(32,32),(16,16)])"
)

if not exist "%SHORTCUT%" (
    powershell -NoProfile -Command "$ws=New-Object -ComObject WScript.Shell; $s=$ws.CreateShortcut($env:SHORTCUT); $s.TargetPath='pythonw'; $s.Arguments=\"`\"$env:APP_SCRIPT`\"\"; $s.WorkingDirectory=$env:APPDIR; $s.IconLocation=$env:ICO+',0'; $s.Save()"
)

pythonw "%APP_SCRIPT%"
