# Code Signing

Die Windows-Datei `DocuBuilder.exe` wird vom Maintainer mit einem eigenen Code-Signing-Zertifikat signiert. Lokal gebaute oder von GitHub Actions erzeugte, nicht nachträglich signierte Dateien sind **nicht** signiert.

## Ablauf eines Releases
1. Tag setzen und pushen: `git tag v1.0.0 && git push origin v1.0.0`
2. GitHub Actions baut `DocuBuilder.exe` (unsigniert) und legt sie ins Release.
3. Maintainer lädt die .exe herunter und legt sie nach `dist\DocuBuilder.exe`.
4. `sign_windows.bat <Fingerprint>` ausführen. Das Skript signiert, prüft die Signatur und erzeugt `DocuBuilder.exe.sha256`.
5. Signierte Dateien im Release ersetzen: `gh release upload v1.0.0 dist\DocuBuilder.exe dist\DocuBuilder.exe.sha256 --clobber`

## Signatur prüfen
Rechtsklick auf die .exe → Eigenschaften → Digitale Signaturen. Zusätzlich liegt jedem Release eine `.sha256`-Prüfsumme bei:
```
certutil -hashfile DocuBuilder.exe SHA256
```
Der Wert muss mit dem Inhalt der `.sha256`-Datei übereinstimmen.

## SmartScreen und Reputation
Auch eine gültige Signatur verhindert die SmartScreen-Meldung bei neuen Dateien nicht sofort. Die Reputation baut sich über Downloads auf. Hilfreich ist:
- immer mit demselben Zertifikat signieren
- den Dateinamen `DocuBuilder.exe` stabil halten (keine Version im Dateinamen)
- nicht unnötig oft neu bauen und veröffentlichen, denn jede neue Datei hat einen neuen Hash
- die Datei über das Portal „Microsoft Security Intelligence" (Submit a file) einreichen
- Nutzern im Release-Text den Hinweis „Weitere Informationen → Trotzdem ausführen" geben (siehe README)

## Optional: SignPath
Die Pipeline (`.github/workflows/build.yml`) kann alternativ über SignPath signieren, sobald die Repository-Variable `SIGNPATH_ORGANIZATION_ID` gesetzt ist. Wird SignPath Foundation genutzt, muss hier und im README der Hinweis ergänzt werden:
„Free code signing provided by SignPath.io, certificate by SignPath Foundation."

## Datenschutz
Dieses Programm überträgt keine Daten an andere Systeme im Netzwerk, außer der Nutzer hat dies ausdrücklich veranlasst (z. B. durch das Öffnen eines hinterlegten Web-Links).
