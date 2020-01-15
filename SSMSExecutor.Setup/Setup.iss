#ifndef AppVersion
#define AppVersion      '2.0.x-alpha'
#endif

#define SrcDir          'Temp\SSMSExecutor'
#define SsmsOutDir      'SSMSExecutor'
#define SsmsPackageGuid '{{a64d9865-b938-4543-bf8f-a553cc4f67f3}'


[Setup]
AppName=SSMS Executor
AppVersion={#AppVersion}
DefaultDirName={pf}\{#SsmsOutDir}
DisableProgramGroupPage=yes
SourceDir={#SrcDir}
OutputDir=Output
OutputBaseFilename=SSMSExecutor-{#AppVersion}

AllowNoIcons=yes
Compression=lzma2
SolidCompression=yes

PrivilegesRequired=admin
DisableReadyPage=yes
DisableReadyMemo=yes

[Files]
Source: "SSMSExecutor.dll"; DestDir: "{app}"
Source: "SSMSExecutor.dll.config"; DestDir: "{app}"
Source: "Microsoft.SqlServer.TransactSql.ScriptDom.dll"; DestDir: "{app}"
Source: "Resources\Command1Package.ico"; DestDir: "{app}\Resources"
Source: "Resources\license.txt"; DestDir: "{app}\Resources"

Source: "SSMSExecutor.pkgdef"; DestDir: "{code:GetOutputDir|2014}"; Check: CanInstall(2014);
Source: "SSMSExecutor.pkgdef"; DestDir: "{code:GetOutputDir|2016}"; Check: CanInstall(2016);
Source: "SSMSExecutor.pkgdef"; DestDir: "{code:GetOutputDir|2017}"; Check: CanInstall(2017);
Source: "SSMSExecutor.pkgdef"; DestDir: "{code:GetOutputDir|2019}"; Check: CanInstall(2019);

Source: "extension.vsixmanifest"; DestDir: "{code:GetOutputDir|2014}"; Check: CanInstall(2014);
Source: "extension.vsixmanifest"; DestDir: "{code:GetOutputDir|2016}"; Check: CanInstall(2016);
Source: "extension.vsixmanifest"; DestDir: "{code:GetOutputDir|2017}"; Check: CanInstall(2017);
Source: "extension.vsixmanifest"; DestDir: "{code:GetOutputDir|2019}"; Check: CanInstall(2019);

[Ini]
Filename: "{code:GetOutputDir|2014}\SSMSExecutor.pkgdef"; Check: CanInstall(2014); Section: "$RootKey$\Packages\{#SsmsPackageGuid}"; Key: """CodeBase"""; String: """{app}\SSMSExecutor.dll""";
Filename: "{code:GetOutputDir|2016}\SSMSExecutor.pkgdef"; Check: CanInstall(2016); Section: "$RootKey$\Packages\{#SsmsPackageGuid}"; Key: """CodeBase"""; String: """{app}\SSMSExecutor.dll""";
Filename: "{code:GetOutputDir|2017}\SSMSExecutor.pkgdef"; Check: CanInstall(2017); Section: "$RootKey$\Packages\{#SsmsPackageGuid}"; Key: """CodeBase"""; String: """{app}\SSMSExecutor.dll""";
Filename: "{code:GetOutputDir|2019}\SSMSExecutor.pkgdef"; Check: CanInstall(2019); Section: "$RootKey$\Packages\{#SsmsPackageGuid}"; Key: """CodeBase"""; String: """{app}\SSMSExecutor.dll""";

[Registry]
Root: HKCU; Subkey: "{code:GetRegistryKey|2014}"; ValueType: dword; ValueName: "SkipLoading"; ValueData: 1; Check: CanInstall(2014);
Root: HKCU; Subkey: "{code:GetRegistryKey|2016}"; ValueType: dword; ValueName: "SkipLoading"; ValueData: 1; Check: CanInstall(2016);
Root: HKCU; Subkey: "{code:GetRegistryKey|2017}"; ValueType: dword; ValueName: "SkipLoading"; ValueData: 1; Check: CanInstall(2017);

[Code]
const
  SSMS_2019 = 2019;
  SSMS_2017 = 2017;
  SSMS_2016 = 2016;
  SSMS_2014 = 2014;

  SSMS_HKEY = 'Software\Microsoft\SQL Server Management Studio\%s';
  SSMS_HKEY_CONFIG = SSMS_HKEY + '_Config';
  

type
  TSsms = record
    Version: Word;
    InternalVersion: String;
    InstallDir: String;
    ExtensionsDir: String;
  end;
  
  TSsmsArr = array of TSsms;

var
  SsmsOptionPage: TInputOptionWizardPage;
  Ssms: TSsmsArr;


function GetSsmsInternalVersion(Version: Integer): String;
begin
  Result:= '';
  if Version = SSMS_2014 then Result:= '12.0'
  else if Version = SSMS_2016 then Result:= '13.0'
  else if Version = SSMS_2017 then Result:= '14.0'
  else if Version = SSMS_2019 then Result:= '18.0_IsoShell';
end;

function GetSsmsInstallDir(const InternalVersion: String; var InstallDir: String): Boolean;
var
  SubKeyName: String;
  T: String;
begin
  Result:= False;
  SubKeyName:= Format(SSMS_HKEY_CONFIG, [InternalVersion]);

  if RegQueryStringValue(HKEY_CURRENT_USER, SubKeyName, 'InstallDir', InstallDir) then
  begin
    if (InternalVersion = '18.0_IsoShell') then
    begin
      InstallDir := InstallDir + '\Common7\IDE';
    end;
  
    if FileExists(InstallDir + '\Ssms.exe') then
    begin
      Result:= True;
    end
  end
end;

function GetSsmsExtensionsDir(const InternalVersion: String; const InstallDir: String): String;
var
  SubKeyName: String;
  ExtensionsDir: String;
begin
  SubKeyName:= Format(SSMS_HKEY_CONFIG + '\Initialization', [InternalVersion]);

  if not RegQueryStringValue(
    HKEY_CURRENT_USER, SubKeyName, 'ApplicationExtensionsFolder', ExtensionsDir) then 
  begin
    ExtensionsDir:= InstallDir + '\Extensions';
  end;

  Result:= ExtensionsDir;
end;

function GetOutputDir(const Version: String): String;
var
  I: Integer;
begin
  Result:= '';

  for I:= Low(Ssms) to High(Ssms) do begin
    if SsmsOptionPage.Values[I] then 
    begin
      if Ssms[I].Version = StrToInt(Version) then
      begin
        Result:= Ssms[I].ExtensionsDir + '\SSMSExecutor';
        Break;
      end
    end
  end
end;

function GetRegistryKey(const Version: String): String;
var
  InternalVersion: String;
begin
  InternalVersion:= GetSsmsInternalVersion(StrToInt(Version));
  Result:= Format(SSMS_HKEY + '\Packages\%s', [InternalVersion, ExpandConstant('{#SsmsPackageGuid}')]);
end;

function CanInstall(Version: Integer): Boolean;
var
  I: Integer;
begin
  Result:= False;

  for I:= Low(Ssms) to High(Ssms) do begin
    if SsmsOptionPage.Values[I] then 
    begin
      if Ssms[I].Version = Version then
      begin
        Result:= True;
        Break;
      end
    end
  end
end;

function InitializeSetup(): Boolean;
var
  I, Len: Word;
  SsmsArr: array of Word;
  InstallDir, Version, InternalVersion: String;
begin
  Result:= True;

  SsmsArr:= [
    SSMS_2014,
    SSMS_2016,
    SSMS_2017,
	SSMS_2019
  ];

  // Loop through all supported SSMS versions
  for I:= Low(SsmsArr) to High(SsmsArr) do
  begin
    InternalVersion:= GetSsmsInternalVersion(SsmsArr[I]);

    // Check if SSMS is actually installed
    if GetSsmsInstallDir(InternalVersion, InstallDir) then
    begin
      Len:= Length(Ssms);
      SetLength(Ssms, Len + 1);
      
      // SSMS istalled, add it to the available list
      Ssms[Len].Version         := SsmsArr[I];
      Ssms[Len].InternalVersion := InternalVersion;
      Ssms[Len].InstallDir      := InstallDir;
      Ssms[Len].ExtensionsDir   := GetSsmsExtensionsDir(InternalVersion, InstallDir);
    end
  end;

  if Length(Ssms) = 0 then
  begin
    MsgBox('SQL Server Management Studio not installed.', mbCriticalError, mb_Ok);
    Result:= False;
  end
end;

procedure InitializeWizard();
var
  LabelText: String;
  Idx: Integer;
begin
  SsmsOptionPage:= CreateInputOptionPage(wpWelcome,
    'Features', 'Select features to install',
    'Please select features to be installed',
    False, False);

  for Idx:= Low(Ssms) to High(Ssms) do
  begin
    LabelText:= Format('SQL Server Management Studio %d Extension', [Ssms[Idx].Version]);
    SsmsOptionPage.Add(LabelText);
    SsmsOptionPage.Values[Idx]:= Idx = High(Ssms);
  end
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  SelectedFeatures, I: Integer;
begin
  Result:= True;

  if CurPageID = 100 then
  begin
    for I:= Low(Ssms) to High(Ssms) do
    begin
      if SsmsOptionPage.Values[I] then Inc(SelectedFeatures);
    end;

    if SelectedFeatures = 0 then
    begin
      MsgBox('Please select at least one SSMS version', mbInformation, MB_OK);
      Result:= False;
    end
  end
end;