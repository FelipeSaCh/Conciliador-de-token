#ifndef Version
  #define Version "0.0.0"
#endif
#define AppNombre "Auditor"
#define AppExe "Auditor.exe"

[Setup]
AppId={{06C528AE-070E-4824-8E6E-CF9C40FD6927}
AppName={#AppNombre}
AppVersion={#Version}
AppVerName={#AppNombre} {#Version}
AppPublisher=Asesores Contables y Tributarios
VersionInfoVersion={#Version}
DefaultDirName={localappdata}\Programs\{#AppNombre}
DefaultGroupName={#AppNombre}
DisableProgramGroupPage=yes
DisableDirPage=yes
PrivilegesRequired=lowest
OutputDir=instaladores
OutputBaseFilename=Auditor_Setup_{#Version}
SetupIconFile=assets\icon.ico
UninstallDisplayIcon={app}\{#AppExe}
Compression=lzma2
SolidCompression=yes
CloseApplications=force
RestartApplications=no
WizardStyle=modern

[Languages]
Name: "es"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "escritorio"; Description: "Crear acceso directo en el escritorio"

[Files]
Source: "dist\main.dist\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[InstallDelete]
Type: filesandordirs; Name: "{app}\*"

[Icons]
Name: "{group}\{#AppNombre}"; Filename: "{app}\{#AppExe}"
Name: "{userdesktop}\{#AppNombre}"; Filename: "{app}\{#AppExe}"; Tasks: escritorio

[Run]
Filename: "{app}\{#AppExe}"; Description: "Abrir {#AppNombre}"; Flags: nowait postinstall
