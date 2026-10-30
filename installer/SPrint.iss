; SPrint installer (Inno Setup 6.5+). Packages the PyInstaller folder build:
;   pyinstaller SPrint.spec
;   iscc /DAppVersion=2.0.0 installer\SPrint.iss      ->  dist\SPrint-Setup-2.0.0.exe
; Optional Croatian wizard texts: put Croatian.isl (Inno Setup's unofficial translation) next to this file.

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#define AppName "SPrint"
#define AppExe "SPrint.exe"
#if FileExists(AddBackslash(SourcePath) + "Croatian.isl")
  #define HaveCroatian
#endif

[Setup]
AppId={{66A52FB3-9494-4CF7-B468-C1D630F072FB}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher=Filip Kozina
AppPublisherURL=https://github.com/Argo8/SPrint
VersionInfoVersion={#AppVersion}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
; Installs for the current user without administrator rights; "all users" can be chosen in the first dialog.
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog commandline
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
CloseApplications=yes
RestartApplications=no
WizardStyle=modern
SetupIconFile=..\icon.ico
UninstallDisplayIcon={app}\{#AppExe}
UninstallDisplayName={#AppName}
OutputDir=..\dist
OutputBaseFilename=SPrint-Setup-{#AppVersion}
Compression=lzma2/max
SolidCompression=yes
ShowLanguageDialog=auto

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"
#ifdef HaveCroatian
Name: "hr"; MessagesFile: "Croatian.isl"
#endif

[CustomMessages]
en.SendToShortcut=Add SPrint to the Explorer "Send to" menu (right-click PDF files or folders)
#ifdef HaveCroatian
hr.SendToShortcut=Dodaj SPrint u izbornik „Pošalji u” (desni klik na PDF datoteke ili mape)
#endif

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "sendto"; Description: "{cm:SendToShortcut}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "..\dist\SPrint\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon
Name: "{usersendto}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: sendto

[Run]
Filename: "{app}\{#AppExe}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent
