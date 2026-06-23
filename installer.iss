; installer.iss — Inno Setup script for TV Pipeline
; Compile with Inno Setup 6:  https://jrsoftware.org/isdl.php
; Or just run build.bat — it compiles this automatically if Inno Setup is installed.

[Setup]
AppName=TV Pipeline
AppVersion=1.0
AppPublisher=Geotechnical Department
; {autopf} picks Program Files or Program Files (x86) based on the exe bitness
DefaultDirName={autopf}\TvPipeline
DefaultGroupName=TV Pipeline
; Output goes into the same dist\ folder as the raw exe
OutputDir=dist
OutputBaseFilename=TvPipeline_Setup
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
; Require admin so we can write to Program Files and set folder permissions
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
; The exe produced by PyInstaller
Source: "dist\TvPipeline.exe"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
; Grant the Users group write access so the app can save config/logs
; without needing admin at runtime
Name: "{app}\config"; Permissions: users-modify
Name: "{app}\logs";   Permissions: users-modify

[Icons]
Name: "{group}\TV Pipeline";           Filename: "{app}\TvPipeline.exe"
Name: "{group}\Uninstall TV Pipeline"; Filename: "{uninstallexe}"
Name: "{commondesktop}\TV Pipeline";   Filename: "{app}\TvPipeline.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\TvPipeline.exe"; Description: "Launch TV Pipeline now"; Flags: nowait postinstall skipifsilent
