[Setup]
AppId={{C6F124E7-8C8B-4B51-9A33-F4A39A3AD7F9}}
AppName=Clipboard To File
AppVersion=1.0.0
AppPublisher=ByronAP
AppPublisherURL=https://github.com/ByronAP/ClipboardToFile
AppSupportURL=https://github.com/ByronAP/ClipboardToFile/issues
AppUpdatesURL=https://github.com/ByronAP/ClipboardToFile/releases
DefaultDirName={autopf}\Clipboard To File
DefaultGroupName=Clipboard To File
DisableProgramGroupPage=yes
OutputBaseFilename=ClipboardToFile-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "x64\Release\ClipboardToFile.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Clipboard To File"; Filename: "{app}\ClipboardToFile.exe"
Name: "{commondesktop}\Clipboard To File"; Filename: "{app}\ClipboardToFile.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\ClipboardToFile.exe"; Description: "{cm:LaunchProgram,Clipboard To File}"; Flags: nowait postinstall skipifsilent

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}";
Name: "startup"; Description: "Start Clipboard To File with Windows"; GroupDescription: "Startup Options"; Flags: checkedonce

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "ClipboardToFile"; ValueData: """{app}\ClipboardToFile.exe"""; Flags: uninsdeletevalue; Tasks: startup