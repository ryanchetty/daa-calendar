; ============================================================
; File: installer.iss
; ============================================================

#define MyAppName "DAACal"
#define MyAppExeName "DAA_Calendar.exe"
#define MyAppPublisher "Ryan Chetty"
#define MyAppVersion "1.2"

[Setup]
AppId={{2DAEDB6B-9B72-4F6D-8B67-3E6B86E2D2B1}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}

DefaultDirName={autopf}\DAACalendar
UsePreviousAppDir=yes

DefaultGroupName=DAACal
DisableProgramGroupPage=yes

OutputDir=output
OutputBaseFilename=DAACal_Installer
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible

; Installer icon (EXE for the installer itself)
SetupIconFile=DAACal.ico

; ARP (Add/Remove Programs) icon
UninstallDisplayIcon={app}\DAACal.ico

; Better update behavior
CloseApplications=yes
RestartApplications=no
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
; Main EXE (onefile)
Source: "dist\DAA_Calendar.exe"; DestDir: "{app}"; DestName: "{#MyAppExeName}"; Flags: ignoreversion

; Install the icon so shortcuts + uninstall icon can reference it
Source: "DAACal.ico"; DestDir: "{app}"; Flags: ignoreversion

; Optional extras you ship with app
; Source: "DAA Claiming Spreadsheet.xlsx"; DestDir: "{app}"; Flags: ignoreversion
; Source: "icons\*"; DestDir: "{app}\icons"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\DAACal"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\DAACal.ico"
Name: "{group}\Uninstall DAACal"; Filename: "{uninstallexe}"; IconFilename: "{app}\DAACal.ico"
Name: "{autodesktop}\DAACal"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; IconFilename: "{app}\DAACal.ico"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch DAACal"; Flags: nowait postinstall skipifsilent