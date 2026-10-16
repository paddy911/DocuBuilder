@echo off
REM Erzeugt dist\DocuBuilder-Setup-<Version>.exe (Installationsprogramm)
REM Benoetigt Python 3 und Inno Setup 6 (winget install JRSoftware.InnoSetup)
setlocal
cd /d "%~dp0"
for /f %%v in ('py tools\version.py show') do set VER=%%v

py -m pip install --upgrade pyinstaller || goto :fehler
py -m PyInstaller --noconfirm --onefile --windowed --name DocuBuilder --version-file version_info.txt --icon assets/docubuilder.ico --add-data "assets/docubuilder.ico:assets" --add-data "assets/docubuilder.png:assets" docubuilder.py || goto :fehler

set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" goto :kein_inno

"%ISCC%" /DAppVersion=%VER% installer\DocuBuilder.iss || goto :fehler
echo.
echo Fertig: dist\DocuBuilder-Setup-%VER%.exe
pause
exit /b 0

:kein_inno
echo Inno Setup 6 nicht gefunden. Installieren mit: winget install JRSoftware.InnoSetup
:fehler
pause
exit /b 1
