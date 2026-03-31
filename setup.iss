; ================================================
;  Mail Merge - Inno Setup Script
;  NON richiede PyInstaller.
;  Installa Python embedded + script direttamente.
;  Richiede: Inno Setup 6+ (https://jrsoftware.org)
; ================================================

#define AppName "Mail Merge"
#define AppVersion "1.0.2"
#define AppPublisher "salvotrotta.com"
#define AppExeName "avvia.bat"

[Setup]
AppId={{B2C3D4E5-F6A7-8901-BCDE-F12345678901}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
SetupIconFile=icon.ico
OutputDir=installer_output
OutputBaseFilename=MailMerge_Setup_v{#AppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
WizardStyle=modern
MinVersion=10.0
Uninstallable=yes
UninstallDisplayName={#AppName}

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"

[Tasks]
Name: "desktopicon"; Description: "Crea un'icona sul Desktop"; GroupDescription: "Icone aggiuntive:"; Flags: unchecked

[Files]
; Script principale
Source: "mail_merge_gui.py"; DestDir: "{app}"; Flags: ignoreversion
; Icona
Source: "icon.ico"; DestDir: "{app}"; Flags: ignoreversion
; Launcher VBS (avvia senza finestra nera)
Source: "avvia.bat"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\avvia.bat"; IconFilename: "{app}\icon.ico"
Name: "{group}\Disinstalla {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\avvia.bat"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon

[Run]
; Installa dipendenze Python alla fine dell'installazione
Filename: "pip"; Parameters: "install python-docx openpyxl --quiet"; \
    Flags: runhidden waituntilterminated; \
    StatusMsg: "Installo dipendenze Python...";
; Avvia l'app
Filename: "{app}\avvia.bat"; Description: "Avvia {#AppName}"; \
    Flags: nowait postinstall skipifsilent shellexec

[Code]
{ Verifica che Python sia installato prima di procedere }
function IsPythonInstalled(): Boolean;
var
  ResultCode: Integer;
begin
  Result := Exec('python', '--version', '', SW_HIDE, ewWaitUntilTerminated, ResultCode)
            and (ResultCode = 0);
end;

function InitializeSetup(): Boolean;
begin
  if not IsPythonInstalled() then
  begin
    MsgBox(
      'Python non e'' installato sul computer.' + #13#10 + #13#10 +
      'Per installare Mail Merge e'' necessario Python 3.8 o superiore.' + #13#10 +
      'Scaricalo da: https://www.python.org/downloads/' + #13#10 + #13#10 +
      'IMPORTANTE: durante l''installazione di Python,' + #13#10 +
      'spunta la casella "Add Python to PATH".',
      mbError, MB_OK
    );
    Result := False;
  end
  else
    Result := True;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
begin
  if CurStep = ssPostInstall then
  begin
    { Installa dipendenze }
    Exec('python', '-m pip install python-docx openpyxl --quiet',
         '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  end;
end;