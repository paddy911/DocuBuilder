@echo off
REM Erzeugt dist\DocuBuilder.exe (portabel, eine einzige Datei, unsigniert)
py -m pip install --upgrade pyinstaller
py -m PyInstaller --onefile --windowed --name DocuBuilder --version-file version_info.txt docubuilder.py
echo.
echo Fertig: dist\DocuBuilder.exe
pause
