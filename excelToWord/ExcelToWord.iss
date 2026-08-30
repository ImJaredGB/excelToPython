#define MyAppName "Generador de documentos"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "JaredGB"
#define MyAppExeName "main.exe"

[Setup]
AppId={{A8B7C6D5-E4F3-42A1-9876-123456789ABC}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}

DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}

OutputDir=installer
OutputBaseFilename=Excel_a_Word_Setup

Compression=lzma
SolidCompression=yes

PrivilegesRequired=admin

WizardStyle=modern

[Files]
Source: "dist\main.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "plantilla_control.docx"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autodesktop}\Excel a Word"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Excel a Word"; Filename: "{app}\{#MyAppExeName}"

[UninstallDelete]
Type: filesandordirs; Name: "{app}"