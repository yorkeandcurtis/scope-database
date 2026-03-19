; Inno Setup Script for Y&C Proposal Builder
; Yorke & Curtis, Inc.

[Setup]
AppName=Y&C Proposal Builder
AppVersion=1.0
AppPublisher=Yorke & Curtis, Inc.
DefaultDirName={autopf}\Proposal Builder
DefaultGroupName=Yorke & Curtis
OutputDir=installer_output
OutputBaseFilename=ProposalBuilder-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
DisableProgramGroupPage=yes
PrivilegesRequired=lowest

[Files]
Source: "dist\Proposal Builder\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\Proposal Builder"; Filename: "{app}\Proposal Builder.exe"
Name: "{autodesktop}\Proposal Builder"; Filename: "{app}\Proposal Builder.exe"

[Run]
Filename: "{app}\Proposal Builder.exe"; Description: "Launch Proposal Builder"; Flags: nowait postinstall skipifsilent
