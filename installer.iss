[Setup]
AppName=SDW Generator
AppVersion=2.0.1
DefaultDirName={autopf}\SDW_Generator
DefaultGroupName=SDW Generator
OutputDir=dist_installer
OutputBaseFilename=SDW_Installer_v2.0.1
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin

; --- BUENAS PRÁCTICAS PARA AUTO-UPDATES Y ROBUSTEZ ---
; 1. Identificador único para detectar si la app está en ejecución
AppMutex=SDW_Generator_Mutex

; 2. Cierra automáticamente instancias previas si están abiertas
CloseApplications=yes
RestartApplications=yes

[Files]
; Archivo ejecutable principal
Source: "dist\SDW_Generator.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\SDW Generator"; Filename: "{app}\SDW_Generator.exe"
Name: "{commondesktop}\SDW Generator"; Filename: "{app}\SDW_Generator.exe"; IconFilename: "{app}\SDW_Generator.exe"

[Run]
; 📌 CORRECCIÓN: Al quitar 'skipifsilent', Inno Setup relanzará la app tanto en modo gráfico como en /VERYSILENT
Filename: "{app}\SDW_Generator.exe"; Description: "Ejecutar aplicación"; Flags: nowait postinstall