# DocuBuilder

Kleines Desktop-Tool für Windows und Linux: Eine Auswahl an Buttons, jeder Button öffnet eine hinterlegte Excel-Liste. Einträge lassen sich im Einstellungsfenster beliebig hinzufügen, ändern, löschen und sortieren.

## Funktionen
- Buttons öffnen Excel-Dateien mit dem Standardprogramm (Windows: Excel, Linux: `xdg-open`)
- Unterstützt lokale Pfade, Netzpfade (`\\Server\Freigabe\datei.xlsx`) und Web-Links
- Einstellungen mit drei Reitern: **Einträge**, **Allgemein** (Ansicht hell/dunkel/System, maximale Anzahl an Einträgen, Zurücksetzen) und **Info**
- Portabel: Einstellungen liegen in `config.json` neben dem Programm
- Keine Abhängigkeiten außer Python 3 (tkinter)

## Starten (ohne Build)
```bash
python docubuilder.py        # Windows
python3 docubuilder.py       # Linux (benötigt python3-tk)
```

## Fertige Programme herunterladen
Unter **Releases** liegen `DocuBuilder.exe` (Windows) und `DocuBuilder` (Linux).

## Selbst bauen
- Windows: `build_windows.bat` doppelklicken, Ergebnis in `dist\`
- Windows signieren: `sign_windows.bat <Fingerprint>` (siehe CODE_SIGNING.md)
- Linux: `./build_linux.sh`, Ergebnis in `dist/`

## Automatischer Build über GitHub
Der Workflow `.github/workflows/build.yml` baut bei jedem Push eines Tags `v*` (z. B. `v1.0.0`) beide Versionen und hängt sie an ein Release.
Zum manuellen Test: Reiter **Actions** → **Build** → **Run workflow**.

## Code Signing
Die Windows-Version wird mit einem eigenen Code-Signing-Zertifikat signiert (`sign_windows.bat`). Ablauf und Prüfung siehe [CODE_SIGNING.md](CODE_SIGNING.md).

## Datenschutz
Dieses Programm überträgt keine Daten an andere Systeme im Netzwerk, außer der Nutzer hat dies ausdrücklich veranlasst.

## Lizenz
MIT
