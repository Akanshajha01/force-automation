; =========================================================
; Force Automation - Inno Setup installer script
;
; Prerequisite: build the app first with build_exe.bat so that
; ..\dist\Force_Automation\Force_Automation.exe exists.
;
; Build the installer with Inno Setup (https://jrsoftware.org/isinfo.php):
;   1. Install Inno Setup on a Windows machine.
;   2. Open this file in Inno Setup Compiler, or run from command line:
;        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
;   3. Output: installer\Output\Force_Automation_Setup.exe
; =========================================================

#define MyAppName "Force Automation"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "CSIR-NPL Force & Hardness Laboratory"
#define MyAppExeName "Force_Automation.exe"

[Setup]
AppId={{8F5B7E3E-4C2A-4D6E-9C1B-FA0F0RCE0001}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
; Not run as admin by default; install location under Program Files
; still requires elevation to install (normal for Windows installers),
; but the app itself never needs admin rights to RUN, since it writes
; certificates/logs to the user's own profile folders.
PrivilegesRequired=lowest
OutputDir=Output
OutputBaseFilename=Force_Automation_Setup
SetupIconFile=app_icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Files]
; Copies the ENTIRE onedir build (exe + all bundled dependencies,
; templates, and assets) into the install folder.
Source: "..\dist\Force_Automation\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
