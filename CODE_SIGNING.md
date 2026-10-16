# Code Signing

Ohne digitale Signatur blockiert **Smart App Control** (Windows 11) DocuBuilder und vor allem das **Update-Setup** („Eine Anwendungssteuerungsrichtlinie hat diese Datei blockiert“, Fehler 4551). SmartScreen warnt zusätzlich bei jedem neuen Download. Deshalb sollen `DocuBuilder.exe` **und** `DocuBuilder-Setup-<Version>.exe` mit einem Zertifikat einer anerkannten Zertifizierungsstelle signiert werden. Ein selbst erstelltes Zertifikat reicht dafür nicht.

## Empfohlen: SignPath (automatisch in GitHub Actions)
Die Pipeline (`.github/workflows/_build.yml`) signiert automatisch, sobald SignPath eingerichtet ist. Solange die Variablen fehlen, baut sie unsigniert weiter.

Für Open-Source-Projekte gibt es kostenloses Signieren über die **SignPath Foundation** (Antrag auf signpath.org). Voraussetzungen erfüllt DocuBuilder bereits: öffentliches Repo, MIT-Lizenz, Build über GitHub Actions.

### Einrichten
1. Bei SignPath (bzw. SignPath Foundation) ein Projekt anlegen und das GitHub-Repo `plange-dev/DocuBuilder` als vertrauenswürdige Build-Quelle verbinden.
2. Zwei **Artifact Configurations** anlegen:
   - für das Programm: Inhalt von `signpath/artifact-configuration.xml` (signiert `DocuBuilder.exe`)
   - für das Setup: Inhalt von `signpath/artifact-configuration-setup.xml` (signiert `DocuBuilder-Setup.exe`)
3. Im GitHub-Repo unter **Settings → Secrets and variables → Actions** eintragen:

   | Art | Name | Inhalt |
   |---|---|---|
   | Secret | `SIGNPATH_API_TOKEN` | API-Token eines SignPath-CI-Benutzers |
   | Variable | `SIGNPATH_ORGANIZATION_ID` | Organisations-ID |
   | Variable | `SIGNPATH_PROJECT_SLUG` | Projekt-Slug |
   | Variable | `SIGNPATH_SIGNING_POLICY_SLUG` | Signing Policy (z. B. `release-signing`) |
   | Variable | `SIGNPATH_ARTIFACT_CONFIGURATION_SLUG` | Slug der Konfiguration für das Programm |
   | Variable | `SIGNPATH_SETUP_ARTIFACT_CONFIGURATION_SLUG` | Slug der Konfiguration für das Setup |

4. Einen Nightly-Build anstoßen (Push auf `nightly` oder **Actions → Nightly → Run workflow**). Im Log stehen „Signaturstatus: Valid“ und „Signaturstatus Setup: Valid“.

Wird die SignPath Foundation genutzt, muss im README und hier der Hinweis stehen:
„Free code signing provided by SignPath.io, certificate by SignPath Foundation.“

### Ablauf
1. PyInstaller baut `DocuBuilder.exe` → SignPath signiert sie.
2. Inno Setup baut das Setup mit der **signierten** `DocuBuilder.exe` → SignPath signiert das Setup.
3. Prüfsummen (`.sha256`) werden erst danach berechnet, passen also zu den signierten Dateien.

### Einschränkung: Deinstallationsprogramm
Das Deinstallationsprogramm (`unins000.exe`) erzeugt Inno Setup selbst; es ist damit **nicht** signiert. Unter Smart App Control kann daher das Deinstallieren blockiert werden. Abhilfe wäre die Inno-Option `SignedUninstaller` mit einem Signierwerkzeug, das während des Kompilierens synchron signiert (z. B. ein lokales Zertifikat oder ein Signierdienst mit Kommandozeilen-Werkzeug).

## Alternative: eigenes Zertifikat lokal
Mit einem gekauften Code-Signing-Zertifikat kann `sign_windows.bat <Fingerprint>` eine heruntergeladene `dist\DocuBuilder.exe` signieren und die Prüfsumme neu erzeugen. Signierte Dateien danach im Release ersetzen, z. B.:
```
gh release upload v1.1.0.0 dist\DocuBuilder.exe dist\DocuBuilder.exe.sha256 --clobber
```
Das Setup muss in diesem Fall ebenfalls signiert und ersetzt werden (mit der bereits signierten `DocuBuilder.exe` neu bauen, dann signieren).

## Signatur prüfen
Rechtsklick auf die .exe → Eigenschaften → Digitale Signaturen, oder in PowerShell:
```
Get-AuthenticodeSignature .\DocuBuilder-Setup-1.0.1.4.exe
```
Zusätzlich liegt jedem Release eine `.sha256`-Prüfsumme bei (`certutil -hashfile <Datei> SHA256`).

## SmartScreen und Reputation
Auch eine gültige Signatur verhindert die SmartScreen-Meldung bei neuen Dateien nicht sofort; die Reputation baut sich über Downloads auf. Hilfreich ist, immer mit demselben Zertifikat zu signieren. Smart App Control dagegen lässt Programme mit gültiger Signatur einer anerkannten Stelle zu.

## Datenschutz
Dieses Programm überträgt keine Daten an andere Systeme im Netzwerk, außer der Nutzer hat dies ausdrücklich veranlasst (z. B. durch das Öffnen eines hinterlegten Web-Links oder die Update-Suche).
