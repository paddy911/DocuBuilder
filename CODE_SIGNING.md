# Code Signing Policy

Free code signing provided by [SignPath.io](https://signpath.io), certificate by [SignPath Foundation](https://signpath.org).

## Was wird signiert?
Nur die Windows-Datei `DocuBuilder.exe`, die in diesem Repository per GitHub Actions (`.github/workflows/build.yml`) aus dem öffentlichen Quellcode gebaut wird. Lokal gebaute Dateien sind nicht signiert.

## Rollen
- Autor / Maintainer / Freigeber: [DEIN GITHUB-NAME](https://github.com/DEIN-GITHUB-NAME)

## Prüfung der Signatur
Rechtsklick auf die .exe → Eigenschaften → Digitale Signaturen. Zusätzlich liegt jedem Release eine `.sha256`-Prüfsumme bei.

## Datenschutz
Dieses Programm überträgt keine Daten an andere Systeme im Netzwerk, außer der Nutzer hat dies ausdrücklich veranlasst (z. B. durch das Öffnen eines hinterlegten Web-Links).
