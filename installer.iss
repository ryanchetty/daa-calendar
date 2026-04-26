; ============================================================
; File: installer.iss
; Build this against a PyInstaller ONEDIR build
; ============================================================

#define MyAppName "DAACal"
#define MyAppExeName "DAA_Calendar.exe"
#define MyAppPublisher "Ryan Chetty"
#define MyAppVersion "2.0.1"
#define MyAppDirName "DAACalendar"
#define MyAppSourceDir "dist\DAA_Calendar"

[Setup]
AppId={{2DAEDB6B-9B72-4F6D-8B67-3E6B86E2D2B1}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}

DefaultDirName={autopf}\{#MyAppDirName}
UsePreviousAppDir=yes

DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes

OutputDir=output
OutputBaseFilename=DAACal_Installer
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=admin

SetupIconFile=DAACal.ico
UninstallDisplayIcon={app}\DAACal.ico

CloseApplications=yes
CloseApplicationsFilter={#MyAppExeName}
RestartApplications=no

WizardStyle=modern
UsePreviousGroup=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "{#MyAppSourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "DAACal.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\DAACal"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\DAACal.ico"
Name: "{group}\Uninstall DAACal"; Filename: "{uninstallexe}"; IconFilename: "{app}\DAACal.ico"
Name: "{autodesktop}\DAACal"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; IconFilename: "{app}\DAACal.ico"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch DAACal"; Flags: nowait postinstall skipifsilent