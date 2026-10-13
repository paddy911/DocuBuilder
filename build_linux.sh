#!/usr/bin/env bash
# Erzeugt dist/DocuBuilder (einzelne Datei). Benötigt python3-tk.
set -e
python3 -m pip install --upgrade pyinstaller
python3 -m PyInstaller --onefile --windowed --name DocuBuilder docubuilder.py
echo "Fertig: dist/DocuBuilder"
