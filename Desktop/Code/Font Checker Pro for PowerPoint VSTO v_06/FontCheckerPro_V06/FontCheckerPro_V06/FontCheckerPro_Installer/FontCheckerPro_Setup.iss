; Backstage Ninja Tools - Inno Setup Script
; Generated for packaging the VSTO add-in

[Setup]
AppName=Backstage Ninja Tools
AppVersion=0.0.7
DefaultDirName={pf}\Backstage Ninja Tools
DefaultGroupName=Backstage Ninja Tools
OutputDir=.
OutputBaseFilename=BackstageNinjaTools_Setup
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64

[Files]
Source: "..\FontCheckerPro_V06\bin\Release\FontCheckerPro_V06.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\FontCheckerPro_V06\bin\Release\FontCheckerPro_V06.dll.manifest"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\FontCheckerPro_V06\bin\Release\FontCheckerPro_V06.vsto"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\FontCheckerPro_V06\bin\Release\FontCheckerPro_V06.xml"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\FontCheckerPro_V06\bin\Release\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\FontCheckerPro_V06\Media Report.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\FontCheckerPro_V06\media_report_32.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\FontCheckerPro_V06\Resources\media_report_32.png"; DestDir: "{app}\Resources"; Flags: ignoreversion
Source: "..\FontCheckerPro_V06\Resources\font_scan_32.png"; DestDir: "{app}\Resources"; Flags: ignoreversion
Source: "..\FontCheckerPro_V06\Resources\Ninja32.png"; DestDir: "{app}\Resources"; Flags: ignoreversion
Source: "..\FontCheckerPro_V06\Resources\Update.png"; DestDir: "{app}\Resources"; Flags: ignoreversion
Source: "..\FontCheckerPro_V06\Resources\Update_Icon.png"; DestDir: "{app}\Resources"; Flags: ignoreversion
Source: "..\FontCheckerPro_V06\Resources\info_16.png"; DestDir: "{app}\Resources"; Flags: ignoreversion

[Icons]
Name: "{group}\Backstage Ninja Tools"; Filename: "{app}\FontCheckerPro_V06.vsto"

[Run]
Filename: "{app}\FontCheckerPro_V06.vsto"; Description: "Launch Backstage Ninja Tools"; Flags: shellexec postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
