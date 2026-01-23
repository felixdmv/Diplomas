; ===================================================================
; ===      Script de Inno Setup para Generador de Diplomas        ===
; ===================================================================

; --- Variables Globales (Fácil de cambiar para futuras versiones) ---
#define MyAppName "Generador de Diplomas UBU"
#define MyAppVersion "1.0.0"  ; <-- DEBE COINCIDIR CON __version__ EN app.py
#define MyAppPublisher "Universidad de Burgos"
#define MyAppURL "https://www.ubu.es"
#define MyAppExeName "GeneradorDiplomasUBU.exe"
#define MyAppExeFolder "GeneradorDiplomasUBU" ; Nombre de la carpeta en 'dist'

[Setup]
; ID único de la aplicación. No lo cambies.
AppId={{D3F4A5B6-C7D8-E9F0-A1B2-C3D4E5F6A7B8}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
; Nombre del archivo de salida del instalador
OutputBaseFilename=Setup_Generador_Diplomas_{#MyAppVersion}
OutputDir=instalador_final
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
; Opción para que el usuario decida si quiere icono en el escritorio
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}";

[Files]
; Aquí le decimos qué archivos meter. Coge TODO lo que hay en la carpeta de PyInstaller
Source: "dist\{#MyAppExeFolder}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Icono en el Menú Inicio
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
; Icono en el Escritorio (si el usuario lo marca)
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Opción para lanzar la app al final de la instalación
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent