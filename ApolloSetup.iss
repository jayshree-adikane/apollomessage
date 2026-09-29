#define MyAppName "Valasys Apollo"
#define MyAppVersion "1.0.0"
#define MyAppExeName "Valasys Apollo.exe"

[Setup]
AppId={{A7E7A4A1-0A8D-4DA8-9B5A-VALASYSAPOLLO}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName={autopf}\Valasys Apollo
DefaultGroupName=Valasys Apollo
OutputDir=installer
OutputBaseFilename=ValasysApolloSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile=valasys.ico
UninstallDisplayIcon={app}\Valasys Apollo.exe

[Files]
Source: "dist\Valasys Apollo.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "valasys.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autodesktop}\Valasys Apollo"; Filename: "{app}\Valasys Apollo.exe"; IconFilename: "{app}\valasys.ico"
Name: "{group}\Valasys Apollo"; Filename: "{app}\Valasys Apollo.exe"; IconFilename: "{app}\valasys.ico"