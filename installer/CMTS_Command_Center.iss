; Inno Setup script for CMTS Command Center
; Build with installer\build_installer.ps1 (or build_installer.bat in the repo root),
; which runs PyInstaller first and passes the version in with /DMyAppVersion=x.y.z.
; Requires Inno Setup 6.3 or newer: https://jrsoftware.org/isinfo.php

#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif

#define MyAppName      "CMTS Command Center"
#define MyAppPublisher "UnarmedWarWolf"
#define MyAppURL       "https://github.com/UnarmedWarWolf/Arris_E6000"
#define MyAppExeName   "CMTS_Command_Center.exe"
#define BuildDir       "..\dist\CMTS_Command_Center"

[Setup]
; AppId identifies this app to Windows. Keep it the same forever so new versions
; upgrade in place instead of installing a second copy.
AppId={{4F179285-DA2E-4602-9D5C-3C2CAA698454}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}/issues
AppUpdatesURL={#MyAppURL}/releases
VersionInfoVersion={#MyAppVersion}
VersionInfoDescription={#MyAppName} Setup

; C:\Program Files\CMTS Command Center by default. The user can change it on the
; "Select Destination Location" page.
DefaultDirName={autopf}\{#MyAppName}
DisableDirPage=no
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes

; 64-bit install into "Program Files" (not "Program Files (x86)")
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0

; Admin install for all users by default; a non-admin can pick "Install for me only",
; which goes to %LOCALAPPDATA%\Programs instead and needs no UAC prompt.
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog

LicenseFile=..\LICENSE
SetupIconFile=..\assets\app.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}
WizardStyle=modern
Compression=lzma2/max
SolidCompression=yes

; If the app is open during an upgrade or uninstall, offer to close it.
CloseApplications=yes
RestartApplications=no

OutputDir=..\dist\installer
OutputBaseFilename=CMTS_Command_Center_Setup_{#MyAppVersion}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
; Checked by default. The user can untick it on the "Select Additional Tasks" page.
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[InstallDelete]
; Clear out the previous build's library folder on upgrade so stale files don't linger.
Type: filesandordirs; Name: "{app}\_internal"

[Files]
Source: "{#BuildDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\LICENSE";    DestDir: "{app}"; DestName: "LICENSE.txt"; Flags: ignoreversion
Source: "..\README.md";  DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}";       Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
// Settings, saved hosts and command history live in %USERPROFILE%\.e6000_command_center,
// not in the install folder, so they survive upgrades. On uninstall, ask whether to
// remove them too (default: keep).
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  DataDir: String;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    DataDir := ExpandConstant('{%USERPROFILE}\.e6000_command_center');
    if DirExists(DataDir) and (not UninstallSilent) then
    begin
      if MsgBox('Also delete your saved settings, hosts and command history?' + #13#10 + #13#10 + DataDir,
                mbConfirmation, MB_YESNO or MB_DEFBUTTON2) = IDYES then
        DelTree(DataDir, True, True, True);
    end;
  end;
end;
