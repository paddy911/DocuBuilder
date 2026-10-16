; Installationsprogramm für DocuBuilder (Inno Setup 6)
; Bauen: build_installer.bat  (oder: ISCC.exe /DAppVersion=1.0.0.0 installer\DocuBuilder.iss)
; Erwartet dist\DocuBuilder.exe (PyInstaller) und erzeugt dist\DocuBuilder-Setup-<Version>.exe.

#ifndef AppVersion
  #define AppVersion "1.0.0.0"
#endif
#ifndef AppChannel
  #define AppChannel "stable"
#endif
#if AppChannel == "nightly"
  #define VerName "DocuBuilder " + AppVersion + " (Nightly)"
#else
  #define VerName "DocuBuilder " + AppVersion
#endif

[Setup]
; AppId nie ändern: daran erkennt Windows Updates und Deinstallation
AppId={{3912EB95-FF31-4D11-AE0E-6D244A30E1CA}
AppName=DocuBuilder
AppVersion={#AppVersion}
AppVerName={#VerName}
AppPublisher=Patrick Lange
AppPublisherURL=https://github.com/paddy911/DocuBuilder
VersionInfoVersion={#AppVersion}
DefaultDirName={autopf}\DocuBuilder
DefaultGroupName=DocuBuilder
DisableProgramGroupPage=yes
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
LicenseFile=..\LICENSE
OutputDir=..\dist
OutputBaseFilename=DocuBuilder-Setup-{#AppVersion}
SetupIconFile=..\assets\docubuilder.ico
UninstallDisplayIcon={app}\DocuBuilder.exe
UninstallDisplayName={#VerName}
CloseApplications=yes
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "..\dist\DocuBuilder.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE"; DestDir: "{app}"; DestName: "LICENSE.txt"; Flags: ignoreversion

[Icons]
; AppUserModelID wie im Programm: angeheftetes Taskleisten-Symbol und Fenster gehören zusammen
Name: "{autoprograms}\DocuBuilder"; Filename: "{app}\DocuBuilder.exe"; AppUserModelID: "PatrickLange.DocuBuilder"
Name: "{autodesktop}\DocuBuilder"; Filename: "{app}\DocuBuilder.exe"; AppUserModelID: "PatrickLange.DocuBuilder"; Tasks: desktopicon

[Run]
; Gemeinsamen Einstellungsordner (C:\ProgramData\DocuBuilder) mit den passenden Rechten anlegen
Filename: "{app}\DocuBuilder.exe"; Parameters: "--setup-machine"; StatusMsg: "Einstellungsordner wird eingerichtet ..."; Flags: runhidden waituntilterminated
; Programm danach als normaler Benutzer starten (nicht mit Adminrechten)
Filename: "{app}\DocuBuilder.exe"; Description: "{cm:LaunchProgram,DocuBuilder}"; Flags: nowait postinstall skipifsilent runasoriginaluser

[Code]
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  Dir: String;
  ResultCode: Integer;
begin
  if CurUninstallStep <> usPostUninstall then
    Exit;
  Dir := ExpandConstant('{commonappdata}\DocuBuilder');
  if not DirExists(Dir) or UninstallSilent then
    Exit;
  if MsgBox('Sollen auch alle Einstellungen gelöscht werden?' + #13#10 +
            '(Einträge, Benutzer und Passwörter für alle Windows-Benutzer)' + #13#10#13#10 +
            'Bei "Nein" bleiben sie für eine spätere Neuinstallation erhalten.',
            mbConfirmation, MB_YESNO or MB_DEFBUTTON2) = IDYES then
  begin
    Exec(ExpandConstant('{cmd}'), '/c attrib -R -H "' + Dir + '\*" & rmdir /s /q "' + Dir + '"',
         '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    DelTree(ExpandConstant('{userappdata}\DocuBuilder'), True, True, True);
  end;
end;
