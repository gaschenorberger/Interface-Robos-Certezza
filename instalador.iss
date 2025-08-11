[Setup]
AppName=Certezza - Robos
AppVersion=1.0.1
AppPublisher=Certezza
DefaultDirName={autopf}\Certezza - Robos
DefaultGroupName=Certezza - Robos
OutputDir=dist_instalador
OutputBaseFilename=Setup_Certezza_Robos
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
SetupIconFile=C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\telaRobos\img\logoICO.ico
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "portuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[LicenseFile]
; Arquivo de licença (você pode criar um .txt com o texto da licença)
LicenseFile=C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\telaRobos\licenca.txt

[Files]
; Inclui o executável
Source: "C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\telaRobos\dist\Certezza - Robos\Certezza - Robos.exe"; DestDir: "{app}"; Flags: ignoreversion

; Inclui a pasta _internal
Source: "C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\telaRobos\dist\Certezza - Robos\_internal\*"; DestDir: "{app}\_internal"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
; Atalho na área de trabalho (para todos os usuários)
Name: "{commondesktop}\Certezza - Robos"; Filename: "{app}\Certezza - Robos.exe"

; Atalho no menu iniciar
Name: "{autoprograms}\Certezza - Robos"; Filename: "{app}\Certezza - Robos.exe"


; Atalho no menu iniciar
Name: "{group}\Certezza - Robos"; Filename: "{app}\Certezza - Robos.exe"

[Run]
Filename: "{app}\Certezza - Robos.exe"; Description: "Executar Certezza - Robos"; Flags: nowait postinstall skipifsilent
