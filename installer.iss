[Setup]
AppName=SDW Generator
AppVersion=1.0
; Se recomienda usar {autopf} para que se adapte a 32/64 bits automáticamente
DefaultDirName={autopf}\SDW_Generator
DefaultGroupName=SDW Generator
OutputDir=dist_installer
OutputBaseFilename=SDW_Installer_v1.0
Compression=lzma
SolidCompression=yes
; Privilegios administrativos para instalar en Program Files
PrivilegesRequired=admin 

[Files]
; IMPORTANTE: Apuntar a DIST, no a BUILD
; Si usaste --onedir (carpeta):
;Source: "dist\SDW_Generator\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

; Si usaste --onefile (un solo .exe), comenta la línea de arriba y usa esta:
Source: "dist\SDW_Generator.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\SDW Generator"; Filename: "{app}\SDW_Generator.exe"
Name: "{commondesktop}\SDW Generator"; Filename: "{app}\SDW_Generator.exe"; IconFilename: "{app}\SDW_Generator.exe"

[Run]
Filename: "{app}\SDW_Generator.exe"; Description: "Ejecutar aplicación"; Flags: nowait postinstall skipifsilent