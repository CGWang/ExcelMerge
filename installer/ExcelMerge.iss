#define MyAppName "ExcelMerge"
#define MyAppVersion "2.0.0"
#define MyAppPublisher "ExcelMerge Contributors"
#define MyAppExeName "ExcelMerge.GUI.exe"
#define PublishDir "..\ExcelMerge.GUI\bin\Release\net8.0-windows\win-x64\publish"

[Setup]
AppId={{B8A3D4E1-5C7F-4A2B-9E6D-1F3C5A7B9D2E}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=output
OutputBaseFilename=ExcelMerge-{#MyAppVersion}-setup
Compression=lzma2/ultra64
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "chinese"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Files]
Source: "{#PublishDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"

[Registry]
; Register as diff tool for .xlsx files
Root: HKLM; Subkey: "SOFTWARE\ExcelMerge"; ValueType: string; ValueName: "InstallPath"; ValueData: "{app}"; Flags: uninsdeletekey

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch ExcelMerge"; Flags: nowait postinstall skipifsilent
