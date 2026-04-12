#define MyAppId "{{A3315470-0B11-4A37-9FA8-70D905A113D4}}"
#define MyAppName "Office Copy as Markdown"
#define MyAppPublisher "office-copy-as-markdown"
#define RequiredWindowsDesktopRuntime "10.0.0"
#define RuntimeDownloadUrl "https://dotnet.microsoft.com/download/dotnet/10.0"

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif

#ifndef SourceDir
  #error SourceDir is required.
#endif

#ifndef OutputDir
  #error OutputDir is required.
#endif

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#AppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL="https://github.com/"
DefaultDirName={localappdata}\Programs\OfficeCopyAsMarkdown
DefaultGroupName={#MyAppName}
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputDir={#OutputDir}
OutputBaseFilename=OfficeCopyAsMarkdown-Setup-{#AppVersion}
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\OfficeCopyAsMarkdown.exe
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UsedUserAreasWarning=no
CloseApplications=yes
CloseApplicationsFilter=OfficeCopyAsMarkdown.exe
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; Flags: unchecked
Name: "launchafterinstall"; Description: "Launch Office Copy as Markdown"; Flags: unchecked

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Office Copy as Markdown"; Filename: "{app}\OfficeCopyAsMarkdown.exe"
Name: "{userdesktop}\Office Copy as Markdown"; Filename: "{app}\OfficeCopyAsMarkdown.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\OfficeCopyAsMarkdown.exe"; Description: "Launch Office Copy as Markdown"; Flags: nowait postinstall skipifsilent; Tasks: launchafterinstall

[Code]
function CloseRunningApp(): Boolean;
var
  ResultCode: Integer;
begin
  Result := Exec(
    ExpandConstant('{cmd}'),
    '/C taskkill /IM OfficeCopyAsMarkdown.exe /T /F',
    '',
    SW_HIDE,
    ewWaitUntilTerminated,
    ResultCode);
end;

function ParseVersionPart(const Value: string): Integer;
begin
  Result := StrToIntDef(Value, 0);
end;

function NextVersionPart(var Version: string): Integer;
var
  SeparatorIndex: Integer;
  Part: string;
begin
  SeparatorIndex := Pos('.', Version);
  if SeparatorIndex = 0 then
  begin
    Part := Version;
    Version := '';
  end
  else
  begin
    Part := Copy(Version, 1, SeparatorIndex - 1);
    Delete(Version, 1, SeparatorIndex);
  end;

  Result := ParseVersionPart(Part);
end;

function CompareVersions(const LeftVersion, RightVersion: string): Integer;
var
  LeftRemaining: string;
  RightRemaining: string;
  Index: Integer;
  LeftPart: Integer;
  RightPart: Integer;
begin
  LeftRemaining := LeftVersion;
  RightRemaining := RightVersion;

  for Index := 0 to 3 do
  begin
    if LeftRemaining <> '' then
      LeftPart := NextVersionPart(LeftRemaining)
    else
      LeftPart := 0;

    if RightRemaining <> '' then
      RightPart := NextVersionPart(RightRemaining)
    else
      RightPart := 0;

    if LeftPart < RightPart then
    begin
      Result := -1;
      exit;
    end;

    if LeftPart > RightPart then
    begin
      Result := 1;
      exit;
    end;
  end;

  Result := 0;
end;

function HasRequiredWindowsDesktopRuntimeInView(const RootKey: Integer): Boolean;
var
  ValueNames: TArrayOfString;
  Index: Integer;
begin
  Result := False;
  if not RegGetValueNames(RootKey, 'SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedfx\Microsoft.WindowsDesktop.App', ValueNames) then
    exit;

  for Index := 0 to GetArrayLength(ValueNames) - 1 do
  begin
    if CompareVersions(ValueNames[Index], '{#RequiredWindowsDesktopRuntime}') >= 0 then
    begin
      Result := True;
      exit;
    end;
  end;
end;

function HasRequiredWindowsDesktopRuntime: Boolean;
begin
  Result :=
    HasRequiredWindowsDesktopRuntimeInView(HKLM32) or
    HasRequiredWindowsDesktopRuntimeInView(HKLM64);
end;

function InitializeSetup(): Boolean;
begin
  if HasRequiredWindowsDesktopRuntime() then
  begin
    Result := True;
    exit;
  end;

  MsgBox(
    'Office Copy as Markdown requires Microsoft Windows Desktop Runtime 10 (x64).' + #13#10 + #13#10 +
    'Install it first, then run this installer again.' + #13#10 +
    '{#RuntimeDownloadUrl}',
    mbCriticalError,
    MB_OK);
  Result := False;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
    CloseRunningApp();
end;
