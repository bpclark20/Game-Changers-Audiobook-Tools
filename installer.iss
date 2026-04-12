#define MyAppId "{{CB8F471A-B9B6-49D6-A975-FC0F9FF7C67F}}"
#define MyAppName "JingleAllTheDay"
#ifndef MyAppVersion
  #define MyAppVersion "1.0.0"
#endif
#define MyAppPublisher "JingleAllTheDay"
#define MyAppExeName "JingleAllTheDay.exe"

#define SourceRoot GetEnv("JATD_SOURCE_DIR")
#if SourceRoot == ""
  #define SourceRoot "dist"
#endif

#ifndef MyOutputDir
  #define MyOutputDir "installer"
#endif

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir={#MyOutputDir}
OutputBaseFilename=JingleAllTheDay-Setup-{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
SetupLogging=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "{#SourceRoot}\\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{autoprograms}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Start {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
var
  RemoveUserData: Boolean;

function QueryUninstallStringFromRoot(const RootKey: Integer; var UninstallCommand: string): Boolean;
var
  KeyPath: string;
begin
  KeyPath :=
    'Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\' +
    ExpandConstant('{#MyAppId}_is1');
  Result := RegQueryStringValue(RootKey, KeyPath, 'QuietUninstallString', UninstallCommand);
  if not Result then
    Result := RegQueryStringValue(RootKey, KeyPath, 'UninstallString', UninstallCommand);
end;

function GetExistingUninstallCommand(var UninstallCommand: string): Boolean;
begin
  Result := QueryUninstallStringFromRoot(HKCU, UninstallCommand);
  if not Result then
    Result := QueryUninstallStringFromRoot(HKLM, UninstallCommand);
end;

function InitializeSetup(): Boolean;
var
  UninstallCommand: string;
  Choice: Integer;
  ExitCode: Integer;
begin
  Result := True;

  if not GetExistingUninstallCommand(UninstallCommand) then
    Exit;

  Choice := SuppressibleMsgBox(
    ExpandConstant('{#MyAppName}') + ' is already installed.' + #13#10#13#10 +
    'Yes: Re-install/upgrade now' + #13#10 +
    'No: Uninstall the current version and stop this installer' + #13#10 +
    'Cancel: Exit without making changes',
    mbConfirmation,
    MB_YESNOCANCEL,
    IDYES
  );

  if Choice = IDCANCEL then
  begin
    Result := False;
    Exit;
  end;

  if Choice = IDNO then
  begin
    if Exec(
      ExpandConstant('{cmd}'),
      '/C ' + UninstallCommand,
      '',
      SW_SHOWNORMAL,
      ewWaitUntilTerminated,
      ExitCode
    ) then
    begin
      SuppressibleMsgBox(
        'The current version uninstall command has been started.' + #13#10 +
        'After uninstall completes, run this installer again to install a fresh copy.',
        mbInformation,
        MB_OK,
        IDOK
      );
    end
    else
    begin
      SuppressibleMsgBox(
        'Unable to launch the existing uninstaller. Canceling setup to avoid conflicts.',
        mbCriticalError,
        MB_OK,
        IDOK
      );
    end;

    Result := False;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  DataPath: string;
begin
  if CurUninstallStep = usUninstall then
  begin
    DataPath := ExpandConstant('{userappdata}\\JingleAllTheDay');
    RemoveUserData :=
      SuppressibleMsgBox(
        'Remove settings and database files too?' + #13#10#13#10 +
        'This deletes:' + #13#10 +
        DataPath + '\\settings.ini' + #13#10 +
        DataPath + '\\jingle-library.json',
        mbConfirmation,
        MB_YESNO,
        IDNO
      ) = IDYES;
  end;

  if (CurUninstallStep = usPostUninstall) and RemoveUserData then
  begin
    DelTree(ExpandConstant('{userappdata}\\JingleAllTheDay'), True, True, True);
    DelTree(ExpandConstant('{localappdata}\\JingleAllTheDay'), True, True, True);
  end;
end;
