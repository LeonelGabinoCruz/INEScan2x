; Instalador de INEScan 2x
[Setup]
AppId={{7B1E5C52-3A64-4F0B-9E1D-1AE5CA42B001}
AppName=INEScan 2x
AppVersion=1.1
AppPublisher=INEScan
DefaultDirName={autopf}\INEScan 2x
DefaultGroupName=INEScan 2x
UninstallDisplayIcon={app}\INEScan2x.exe
OutputDir=salida
OutputBaseFilename=Instalar_INEScan2x
SetupIconFile=..\icono.ico
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=yes

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear un acceso directo en el &Escritorio"; GroupDescription: "Accesos directos:"

[Files]
Source: "INEScan2x.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\INEScan 2x"; Filename: "{app}\INEScan2x.exe"
Name: "{autodesktop}\INEScan 2x"; Filename: "{app}\INEScan2x.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\INEScan2x.exe"; Description: "Abrir INEScan 2x ahora"; Flags: nowait postinstall skipifsilent
