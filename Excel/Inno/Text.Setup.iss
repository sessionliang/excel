; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "Text Demo"
#define MyAppVersion "1.0"
#define MyAppPublisher "DM, Inc."
#define MyAppURL "http://sessionliang.github.io/"
#define MyAppExeName "Text.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{BA8DA3DC-0FAD-40D6-A197-316E3960E37D}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName=SiteServer.Service\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=D:\program\excel\test.excel\Excel\Inno
OutputBaseFilename=setupText
SetupIconFile=D:\program\excel\test.excel\Excel\Text\favicon.ico
Compression=lzma
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "D:\program\excel\test.excel\Excel\Text\bin\Debug\Text.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\program\excel\test.excel\Excel\Text\bin\Debug\Text.exe.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\program\excel\test.excel\Excel\Text\bin\Debug\wait.gif"; DestDir: "{app}"; Flags: ignoreversion

; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
