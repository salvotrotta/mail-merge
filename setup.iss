; ================================================
;  Mail Merge - Inno Setup Script
;  Versione standalone (PyInstaller onedir)
;  Non richiede Python installato sul PC.
;  Richiede: Inno Setup 6+ (https://jrsoftware.org)
; ================================================

#define AppName "Mail Merge"
#define AppVersion "1.1.0"
#define AppPublisher "SALVOTROTTA.COM"
#define AppExeName "MailMerge.exe"

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
; Tutto il contenuto della cartella PyInstaller
Source: "dist\MailMerge\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\{#AppExeName}"
Name: "{group}\Disinstalla {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Avvia {#AppName}"; \
    Flags: nowait postinstall skipifsilent
