#define AppName "WorkbookLens"
#define AppPublisher "WorkbookLens contributors"
#define AppURL "https://github.com/chenweixin123/workbooklens"
#define AppExeName "WorkbookLens.exe"

#ifndef AppVersion
  #error AppVersion must be supplied with /DAppVersion=...
#endif

#ifndef NumericVersion
  #error NumericVersion must be supplied with /DNumericVersion=...
#endif

#ifndef PortableRoot
  #error PortableRoot must be supplied with /DPortableRoot=...
#endif

#ifndef OutputDir
  #error OutputDir must be supplied with /DOutputDir=...
#endif

#ifndef OutputBaseFilename
  #error OutputBaseFilename must be supplied with /DOutputBaseFilename=...
#endif

[Setup]
AppId={{7B7534E0-8485-4F4F-8DE7-561869FF7C0C}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}/issues
AppUpdatesURL={#AppURL}/releases
DefaultDirName={localappdata}\Programs\{#AppName}
DefaultGroupName={#AppName}
DisableDirPage=yes
DisableProgramGroupPage=yes
LicenseFile={#PortableRoot}\LICENSE
OutputDir={#OutputDir}
OutputBaseFilename={#OutputBaseFilename}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
Uninstallable=yes
UninstallDisplayName={#AppName}
UninstallDisplayIcon={app}\{#AppExeName}
VersionInfoCompany={#AppPublisher}
VersionInfoDescription={#AppName} local spreadsheet auditor installer
VersionInfoProductName={#AppName}
VersionInfoProductVersion={#AppVersion}
VersionInfoVersion={#NumericVersion}
CloseApplications=yes
RestartApplications=no
SetupLogging=yes
UsePreviousAppDir=no
UsePreviousGroup=yes
UsePreviousTasks=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: checkedonce

[Files]
Source: "{#PortableRoot}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[InstallDelete]
Type: filesandordirs; Name: "{app}\_internal"; Check: ShouldCleanPreviousPayload
Type: filesandordirs; Name: "{app}\LICENSES"; Check: ShouldCleanPreviousPayload
Type: files; Name: "{app}\LICENSE"; Check: ShouldCleanPreviousPayload
Type: files; Name: "{app}\README-PORTABLE.txt"; Check: ShouldCleanPreviousPayload
Type: files; Name: "{app}\Start-WorkbookLens.cmd"; Check: ShouldCleanPreviousPayload
Type: files; Name: "{app}\THIRD-PARTY-NOTICES.txt"; Check: ShouldCleanPreviousPayload
Type: files; Name: "{app}\WorkbookLens.exe"; Check: ShouldCleanPreviousPayload
Type: files; Name: "{app}\workbooklens.example.yml"; Check: ShouldCleanPreviousPayload

[UninstallDelete]
Type: files; Name: "{app}\.workbooklens-install-owner"

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Parameters: "serve --open-browser --fallback-port"; WorkingDir: "{app}"; Comment: "Open the local WorkbookLens interface"; AppUserModelID: "WorkbookLens.Local"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Parameters: "serve --open-browser --fallback-port"; WorkingDir: "{app}"; Comment: "Open the local WorkbookLens interface"; AppUserModelID: "WorkbookLens.Local"; Tasks: desktopicon
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"; WorkingDir: "{app}"

[Run]
Filename: "{app}\{#AppExeName}"; Parameters: "serve --open-browser --fallback-port"; WorkingDir: "{app}"; Description: "{cm:LaunchProgram,{#StringChange(AppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
const
  OwnershipMarkerName = '.workbooklens-install-owner';
  OwnershipMarkerContent =
    'WorkbookLens|{7B7534E0-8485-4F4F-8DE7-561869FF7C0C}|owner-schema=1';
  FileAttributeDirectory = $10;
  FileAttributeReparsePoint = $400;
  InvalidFileAttributes = $FFFFFFFF;

var
  CleanPreviousPayload: Boolean;

function GetFileAttributes(FileName: String): LongWord;
  external 'GetFileAttributesW@kernel32.dll stdcall';

function NormalizePath(Path: String): String;
begin
  Result := RemoveBackslashUnlessRoot(Trim(Path));
end;

function SamePath(LeftPath: String; RightPath: String): Boolean;
begin
  Result :=
    CompareText(NormalizePath(LeftPath), NormalizePath(RightPath)) = 0;
end;

function DefaultInstallDir(): String;
begin
  Result := NormalizePath(
    ExpandConstant('{localappdata}\Programs\{#AppName}'));
end;

function OwnershipMarkerPath(Directory: String): String;
begin
  Result := AddBackslash(NormalizePath(Directory)) + OwnershipMarkerName;
end;

function IsReparsePoint(Path: String): Boolean;
var
  Attributes: LongWord;
begin
  Attributes := GetFileAttributes(Path);
  Result :=
    (Attributes <> InvalidFileAttributes) and
    ((Attributes and FileAttributeReparsePoint) <> 0);
end;

function TreeHasReparsePoint(Path: String): Boolean;
var
  FindRec: TFindRec;
  ChildPath: String;
begin
  Result := False;
  if not (FileExists(Path) or DirExists(Path)) then
    Exit;

  if IsReparsePoint(Path) then
  begin
    Result := True;
    Exit;
  end;

  if not DirExists(Path) then
    Exit;

  if FindFirst(AddBackslash(Path) + '*', FindRec) then
  begin
    try
      repeat
        if (FindRec.Name <> '.') and (FindRec.Name <> '..') then
        begin
          ChildPath := AddBackslash(Path) + FindRec.Name;
          if IsReparsePoint(ChildPath) then
          begin
            Result := True;
            Exit;
          end;
          if ((FindRec.Attributes and FileAttributeDirectory) <> 0) and
             TreeHasReparsePoint(ChildPath) then
          begin
            Result := True;
            Exit;
          end;
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;
end;

function HasValidOwnershipMarker(Directory: String): Boolean;
var
  MarkerContent: AnsiString;
begin
  MarkerContent := '';
  Result :=
    LoadStringFromFile(OwnershipMarkerPath(Directory), MarkerContent) and
    (MarkerContent = OwnershipMarkerContent);
end;

function DirectoryHasEntries(Directory: String): Boolean;
var
  FindRec: TFindRec;
begin
  Result := False;
  if not DirExists(Directory) then
    Exit;

  if FindFirst(AddBackslash(Directory) + '*', FindRec) then
  begin
    try
      repeat
        if (FindRec.Name <> '.') and (FindRec.Name <> '..') then
        begin
          Result := True;
          Exit;
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;
end;

function ValidateInstallTarget(): String;
var
  TargetDir: String;
  ExpectedInstallDir: String;
  InstallParent: String;
begin
  Result := '';
  CleanPreviousPayload := False;
  TargetDir := NormalizePath(ExpandConstant('{app}'));
  ExpectedInstallDir := DefaultInstallDir();
  InstallParent := NormalizePath(ExtractFileDir(ExpectedInstallDir));

  if not SamePath(TargetDir, ExpectedInstallDir) then
  begin
    Result :=
      'WorkbookLens must use the per-user directory "' +
      ExpectedInstallDir + '".';
    Exit;
  end;

  if IsReparsePoint(InstallParent) then
  begin
    Result :=
      'WorkbookLens cannot install because its parent directory is a reparse point: "' +
      InstallParent + '".';
    Exit;
  end;

  if FileExists(TargetDir) then
  begin
    Result :=
      'WorkbookLens cannot install because the target path is a file: "' +
      TargetDir + '".';
    Exit;
  end;

  if not DirectoryHasEntries(TargetDir) then
    Exit;

  if TreeHasReparsePoint(TargetDir) then
  begin
    Result :=
      'WorkbookLens cannot upgrade an installation tree that contains a reparse point: "' +
      TargetDir + '".';
    Exit;
  end;

  if not HasValidOwnershipMarker(TargetDir) then
  begin
    Result :=
      'WorkbookLens cannot upgrade the non-empty directory "' + TargetDir +
      '" because its installer ownership marker is missing or invalid. ' +
      'Move any user files elsewhere and uninstall the older or incomplete installation first.';
    Exit;
  end;

  if not FileExists(AddBackslash(TargetDir) + '{#AppExeName}') then
  begin
    Result :=
      'WorkbookLens cannot upgrade because the owned installation is missing {#AppExeName}.';
    Exit;
  end;

  CleanPreviousPayload := True;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := ValidateInstallTarget();
end;

function ShouldCleanPreviousPayload(): Boolean;
begin
  Result :=
    CleanPreviousPayload and
    SamePath(ExpandConstant('{app}'), DefaultInstallDir()) and
    HasValidOwnershipMarker(ExpandConstant('{app}')) and
    not TreeHasReparsePoint(ExpandConstant('{app}'));
  if CleanPreviousPayload and not Result then
    RaiseException(
      'WorkbookLens stopped because the validated installation tree changed before cleanup.');
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    if not SaveStringToFile(
      OwnershipMarkerPath(ExpandConstant('{app}')),
      OwnershipMarkerContent,
      False) then
      RaiseException('WorkbookLens could not write its installer ownership marker.');
  end;
end;
