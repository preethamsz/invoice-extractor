[Setup]
AppName=Invoice Extractor
AppVersion=2.0
AppPublisher=Chandan Logics
DefaultDirName={autopf}\InvoiceExtractor
DefaultGroupName=Invoice Extractor
OutputDir=C:\project for work\installer
OutputBaseFilename=InvoiceExtractor_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\InvoiceExtractor.exe

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "C:\project for work\dist\InvoiceExtractor.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\project for work\templates\*"; DestDir: "{app}\templates"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\project for work\static\*"; DestDir: "{app}\static"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Invoice Extractor"; Filename: "{app}\InvoiceExtractor.exe"
Name: "{group}\Uninstall Invoice Extractor"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Invoice Extractor"; Filename: "{app}\InvoiceExtractor.exe"
Name: "{commonstartmenu}\Invoice Extractor"; Filename: "{app}\InvoiceExtractor.exe"

[Run]
Filename: "{app}\InvoiceExtractor.exe"; Description: "Launch Invoice Extractor"; Flags: nowait postinstall skipifsilent