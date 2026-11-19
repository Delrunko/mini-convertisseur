#define MyAppName "Mini Convertisseur"
#define MyAppVersion "1.1.0"

[Setup]
AppId={{3E7760E2-4C5A-4C8D-9AE0-2B2DF6BD8E67}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=Delrunko
DefaultDirName={localappdata}\Programs\Mini Convertisseur
DefaultGroupName={#MyAppName}
UninstallDisplayIcon={app}\MiniConvertisseur.exe
OutputDir=installer
OutputBaseFilename=MiniConvertisseur-Setup
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=lowest
WizardStyle=modern
Compression=lzma2
SolidCompression=yes

[Tasks]
Name: "desktopicon"; Description: "Créer une icône sur le Bureau"; GroupDescription: "Options supplémentaires :"; Flags: unchecked

[Files]
Source: "dist\MiniConvertisseur.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\MiniConvertisseur.exe"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\MiniConvertisseur.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\MiniConvertisseur.exe"; Description: "Lancer Mini Convertisseur"; Flags: nowait postinstall skipifsilent
