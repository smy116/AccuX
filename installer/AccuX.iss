; AccuX V1 安装脚本（规格 §26 安装与部署约定）
;
; 编译：ISCC.exe /DAccuXVersion=1.6.1 /DAccuXFileVersion=1.6.1.0 installer\AccuX.iss
; 产物：AccuXSetup-{AccuXVersion}.exe
;
; 要求：
;   - 检查 .NET Framework 4.8 前置条件；缺失时自动下载并安装；
;   - 安装 AccuX 程序集与资源；
;   - 注册 Excel / WPS 所需 COM Add-in 信息（与开发期 register.ps1 同一逻辑来源）；
;   - 根据支持矩阵处理 x86 / x64；
;   - 支持正常卸载；
;   - 为后续版本升级保留稳定的 Product/Upgrade 策略。
;
; 说明：
;   - Inno Setup 官方安装包不自带 ChineseSimplified.isl，已随本工程放在
;     installer\Languages\ 下；升级 Inno Setup 后请同步更新该文件。
;   - Add-in 程序集为 AnyCPU（MSIL），同一份 DLL 由 32 位与 64 位 Office 共用，
;     因此注册表项需要同时写入 32 位与 64 位视图。
;   - 版本号唯一来源是 tag / CI 参数，编译时必须通过 /D 传入；默认值只是
;     开发占位，避免未指定版本时误产出看似正式版的安装包。

#ifndef AccuXVersion
#define AccuXVersion "0.0.0-dev"
#endif
#ifndef AccuXFileVersion
#define AccuXFileVersion "0.0.0.0"
#endif
#ifndef AccuXOutputBaseFilename
#define AccuXOutputBaseFilename "AccuXSetup-" + AccuXVersion
#endif
#define AccuXProgId "AccuX.AddIn.Connect"
#define AccuXFriendlyName "AccuX"
#define SourceRoot "..\src"

; Excel Add-ins 子键名必须是 ProgId：宿主按子键名做 CoCreateInstance。
#define ExcelAddinKey "Software\Microsoft\Office\Excel\Addins\" + AccuXProgId
#define WpsAddinKey "Software\Kingsoft\Office\ET\AddinsWL\" + AccuXProgId

[Setup]
AppId={{8E3B2A64-1C7D-4A9F-9E5B-2D6F0A8C4B31}
AppName=AccuX
AppVersion={#AccuXVersion}
AppVerName=AccuX {#AccuXVersion}
VersionInfoVersion={#AccuXFileVersion}
AppPublisher=AccuX
DefaultDirName={autopf}\AccuX
DefaultGroupName=AccuX
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename={#AccuXOutputBaseFilename}
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=admin
MinVersion=6.1sp1
; 64 位 Windows 上使用原生 64 位 Program Files 与 64 位注册表视图；
; 32 位 Windows 上自动回退为 32 位安装模式（不设 ArchitecturesAllowed）。
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
UninstallDisplayName=AccuX
SetupLogging=yes

[Languages]
; 先读官方默认消息，再由简体中文覆盖，保证个别未翻译消息仍有英文兜底。
Name: "chinese"; MessagesFile: "compiler:Default.isl,Languages\ChineseSimplified.isl"

[Files]
; AccuX 程序集与随附的 Office PIA（Private 部署，不依赖目标机 GAC 布局）
Source: "{#SourceRoot}\AccuX.AddIn\bin\Release\net48\AccuX.AddIn.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\AccuX.AddIn\bin\Release\net48\AccuX.Core.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\AccuX.AddIn\bin\Release\net48\AccuX.Host.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\AccuX.AddIn\bin\Release\net48\AccuX.Modules.BasicFinance.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\AccuX.AddIn\bin\Release\net48\AccuX.Modules.Mark.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\AccuX.AddIn\bin\Release\net48\AccuX.Modules.Compare.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\AccuX.AddIn\bin\Release\net48\Newtonsoft.Json.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\AccuX.AddIn\bin\Release\net48\Microsoft.Office.Interop.Excel.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\AccuX.AddIn\bin\Release\net48\office.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\AccuX.AddIn\bin\Release\net48\Microsoft.Vbe.Interop.dll"; DestDir: "{app}"; Flags: ignoreversion
; 所有文件落盘后，由第一个 [Registry] 条目的代码常量注册 COM。
Source: "{#SourceRoot}\AccuX.AddIn\config.sample.json"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; Excel Add-in 注册（机器级，安装包需要管理员权限）。
; 同一组值分别写入默认视图（64 位安装模式下即 64 位视图）与 32 位视图，
; 保证 64 位 Office 与 32 位 Office 都能读到。
; 首项通过代码常量注册 COM：常量求值失败会中止原生安装事务。
; 不可改为 BeforeInstall/AfterInstall，Inno 会捕获这些回调的异常后继续。
Root: HKLM;   Subkey: "{code:RegisterComAndGetKey|{#ExcelAddinKey}}"; ValueType: string; ValueName: "Description";  ValueData: "AccuX 财务效率插件"; Flags: uninsdeletekey
Root: HKLM;   Subkey: "{#ExcelAddinKey}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AccuXFriendlyName}"
Root: HKLM;   Subkey: "{#ExcelAddinKey}"; ValueType: dword;  ValueName: "LoadBehavior"; ValueData: "3"
Root: HKLM32; Subkey: "{#ExcelAddinKey}"; ValueType: string; ValueName: "Description";  ValueData: "AccuX 财务效率插件"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "{#ExcelAddinKey}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AccuXFriendlyName}"
Root: HKLM32; Subkey: "{#ExcelAddinKey}"; ValueType: dword;  ValueName: "LoadBehavior"; ValueData: "3"

; WPS 表格 Add-in 注册（同样写入两个注册表视图）。
; WPS 的 AddinsWL 是 ProgId 白名单，ProgId 必须是根键下的字符串值，而不是子键。
Root: HKLM;   Subkey: "Software\Kingsoft\Office\ET\AddinsWL"; ValueType: string; ValueName: "{#AccuXProgId}"; ValueData: ""; Flags: uninsdeletevalue
Root: HKLM32; Subkey: "Software\Kingsoft\Office\ET\AddinsWL"; ValueType: string; ValueName: "{#AccuXProgId}"; ValueData: ""; Flags: uninsdeletevalue
Root: HKLM;   Subkey: "{#WpsAddinKey}"; ValueType: string; ValueName: "Description";  ValueData: "AccuX 财务效率插件"; Flags: uninsdeletekey
Root: HKLM;   Subkey: "{#WpsAddinKey}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AccuXFriendlyName}"
Root: HKLM;   Subkey: "{#WpsAddinKey}"; ValueType: dword;  ValueName: "LoadBehavior"; ValueData: "3"
Root: HKLM32; Subkey: "{#WpsAddinKey}"; ValueType: string; ValueName: "Description";  ValueData: "AccuX 财务效率插件"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "{#WpsAddinKey}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AccuXFriendlyName}"
Root: HKLM32; Subkey: "{#WpsAddinKey}"; ValueType: dword;  ValueName: "LoadBehavior"; ValueData: "3"

[UninstallRun]
; 注销 COM 可见程序集（与 [Code] 中注册逻辑对应）。
Filename: "{dotnet4032}\RegAsm.exe"; Parameters: "/unregister ""{app}\AccuX.AddIn.dll"""; Flags: runhidden waituntilterminated logoutput; RunOnceId: "UnregAsm32"
Filename: "{dotnet4064}\RegAsm.exe"; Parameters: "/unregister ""{app}\AccuX.AddIn.dll"""; Flags: runhidden waituntilterminated logoutput; RunOnceId: "UnregAsm64"; Check: IsWin64

[Code]
const
  { .NET Framework 4.8 的 Release 值：
    4.8 = 528040（Win10 1903/1909）、528049（其他系统）、528372（Win10 2004/20H2）；
    4.8.1 = 533325 起。按微软建议使用 >= 比较。 }
  DotNet48MinRelease = 528040;

  { .NET Framework 4.8 Web 安装程序（ndp48-web.exe，约 1.5 MB）。
    fwlink 会 302 重定向到 download.visualstudio.microsoft.com，Inno 下载函数自动跟随。
    刻意不校验 SHA-256：微软重发该引导程序会改变文件哈希，固定哈希会导致安装硬失败；
    完整性依赖 HTTPS 传输与微软官方 fwlink 地址。 }
  Ndp48Url = 'https://go.microsoft.com/fwlink/?LinkId=2085155';
  Ndp48BaseName = 'ndp48-web.exe';

type
  TFileBackup = record
    Destination: String;
    Backup: String;
  end;
  TRegistryBackup = record
    Root: Integer;
    Subkey: String;
    RegExe: String;
    Backup: String;
    Existed: Boolean;
  end;

var
  DownloadPage: TDownloadWizardPage;
  DotNet48RebootRequired: Boolean;
  PrerequisiteError: String;
  FileBackups: array of TFileBackup;
  RegistryBackups: array of TRegistryBackup;
  ComRegistrationStarted: Boolean;
  ComRegistrationComplete: Boolean;
  InstallCommitted: Boolean;

procedure ReportInstallError(const MessageText: String);
begin
  Log(MessageText);
  MsgBox(MessageText, mbCriticalError, MB_OK);
end;

procedure FailInstall(const MessageText: String);
begin
  Log(MessageText);
  { 仅从 ssInstall 或 [Registry] 代码常量调用：异常在这两个位置
    属于安装失败；BeforeInstall/AfterInstall 异常会被捕获后继续。 }
  RaiseException(MessageText);
end;

procedure BackupPackageFile(const Name: String);
var
  Destination, Backup: String;
  I: Integer;
begin
  Destination := ExpandConstant('{app}\') + Name;
  for I := 0 to GetArrayLength(FileBackups) - 1 do
    if CompareText(FileBackups[I].Destination, Destination) = 0 then
      Exit;
  if not FileExists(Destination) then
    Exit;
  I := GetArrayLength(FileBackups);
  Backup := ExpandConstant('{tmp}\accux-file-') + IntToStr(I) + '.bak';
  if not CopyFile(Destination, Backup, True) then
    FailInstall('无法备份旧文件，安装已中止：' + Destination);
  SetArrayLength(FileBackups, I + 1);
  FileBackups[I].Destination := Destination;
  FileBackups[I].Backup := Backup;
end;

procedure BackupRegistryKey(Root: Integer; const RegExe, Subkey: String);
var
  I, ResultCode: Integer;
  Snapshot: TRegistryBackup;
begin
  Snapshot.Root := Root;
  Snapshot.Subkey := Subkey;
  Snapshot.RegExe := RegExe;
  Snapshot.Existed := RegKeyExists(Root, Subkey);
  I := GetArrayLength(RegistryBackups);
  Snapshot.Backup := ExpandConstant('{tmp}\accux-reg-') + IntToStr(I) + '.reg';
  if Snapshot.Existed then
  begin
    if not Exec(RegExe, 'export "HKLM\' + Subkey + '" "' +
        Snapshot.Backup + '" /y', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
      FailInstall('无法启动注册表备份：' + SysErrorMessage(ResultCode));
    if (ResultCode <> 0) or not FileExists(Snapshot.Backup) then
      FailInstall(Format('注册表备份失败（%d）：%s', [ResultCode, Subkey]));
  end;
  SetArrayLength(RegistryBackups, I + 1);
  RegistryBackups[I] := Snapshot;
end;

procedure BackupComView(Root: Integer; const RegExe: String);
begin
  { 当前程序集 RegAsm /regfile 的完整根键集合，包含两个公开枚举 Record。
    修改 COM 可见类型时，必须同步校验此清单（见 installer/README.md）。 }
  BackupRegistryKey(Root, RegExe, 'Software\Classes\{#AccuXProgId}');
  BackupRegistryKey(Root, RegExe, 'Software\Classes\CLSID\{7C1F0E4A-9B2D-4E8C-A1F3-6D5E8B0C2A11}');
  BackupRegistryKey(Root, RegExe, 'Software\Classes\Record\{0513544E-578A-3CEC-B4FF-26A0B3C8F5C1}');
  BackupRegistryKey(Root, RegExe, 'Software\Classes\Record\{118B8701-9F8D-328C-B588-4AD2114FA50C}');
end;

{ 读取 Release 值判断是否已安装 .NET Framework 4.8 或更高版本。 }
function IsDotNet48OrLater(): Boolean;
var
  Release: Cardinal;
begin
  Result := False;
  if RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release) then
    Result := Release >= DotNet48MinRelease;
end;

procedure InitializeWizard;
begin
  DownloadPage := CreateDownloadPage(
    SetupMessage(msgWizardPreparing),
    SetupMessage(msgPreparingDesc),
    nil);
  DownloadPage.ShowBaseNameInsteadOfUrl := True;
end;

function HandleDotNet48ExitCode(ResultCode: Integer): Boolean;
begin
  Result := False;
  PrerequisiteError := '';
  { 0 = 成功；3010 / 1641 = 成功但需重启；1602 = 用户取消；1603 = 致命错误；
    5100 = 不满足系统要求。 }
  Log(Format('.NET Framework 安装程序退出码：%d', [ResultCode]));
  case ResultCode of
    3010, 1641:
      begin
        { 重启返回码优先于 Release 检测：注册表已更新不代表运行时已就绪。 }
        DotNet48RebootRequired := True;
        Result := True;
      end;
    0:
      begin
        if IsDotNet48OrLater() then
          Result := True
        else
          PrerequisiteError := '.NET 安装程序返回成功，但仍未检测到 .NET Framework 4.8。请检查安装日志后重试。';
      end;
    1602:
      PrerequisiteError := '.NET Framework 4.8 安装已被取消。';
    1603:
      PrerequisiteError := '.NET Framework 4.8 安装失败（错误 1603：安装过程中发生致命错误）。' + #13#10#13#10 +
        '请查看安装日志，或手动安装后重新运行本安装程序。';
    5100:
      PrerequisiteError := '.NET Framework 4.8 安装失败（错误 5100：本机不满足系统要求）。';
  else
    PrerequisiteError := Format('.NET Framework 4.8 安装程序返回错误代码 %d。', [ResultCode]);
  end;
end;

{ 下载并安装 .NET Framework 4.8。返回 True 表示已就绪或已交由重启完成。 }
function InstallDotNet48(): Boolean;
var
  InstallerPath: String;
  ResultCode: Integer;
begin
  Result := False;
  PrerequisiteError := '';

  DownloadPage.Clear;
  DownloadPage.Add(Ndp48Url, Ndp48BaseName, '');
  DownloadPage.Show;
  try
    try
      DownloadPage.Download;
    except
      if DownloadPage.AbortedByUser then
        PrerequisiteError := '已取消下载 .NET Framework 4.8 安装程序。'
      else
      begin
        PrerequisiteError := Format('下载 .NET Framework 4.8 安装程序失败：%s', [GetExceptionMessage]);
        PrerequisiteError := PrerequisiteError + #13#10#13#10 +
          '请检查网络连接后重试，或从以下地址手动安装后重新运行本安装程序：' + #13#10 +
          'https://dotnet.microsoft.com/download/dotnet-framework/net48';
      end;
      Exit;
    end;
  finally
    DownloadPage.Hide;
  end;

  InstallerPath := ExpandConstant('{tmp}\' + Ndp48BaseName);
  if not FileExists(InstallerPath) then
  begin
    PrerequisiteError := '未找到已下载的 .NET Framework 4.8 安装程序。';
    Exit;
  end;

  if not Exec(InstallerPath, '/norestart /ChainingPackage AccuX', '',
      SW_SHOW, ewWaitUntilTerminated, ResultCode) then
  begin
    PrerequisiteError := '无法启动 .NET Framework 4.8 安装程序：' + SysErrorMessage(ResultCode);
    Exit;
  end;

  Result := HandleDotNet48ExitCode(ResultCode);
end;

{ 前置条件集中在 PrepareToInstall，统一由常规安装向导处理。 }
function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  NeedsRestart := False;
  if not DotNet48RebootRequired and not IsDotNet48OrLater() then
  begin
    if MsgBox('AccuX 需要 .NET Framework 4.8 或更高版本，当前系统未检测到。' + #13#10#13#10 +
        '是否立即联网下载并安装？（引导程序约 1.5 MB，后续仍需下载运行时组件）',
        mbConfirmation, MB_YESNO) = IDNO then
      Result := '未安装 .NET Framework 4.8，无法继续安装 AccuX。'
    else if not InstallDotNet48() then
      Result := PrerequisiteError;
  end;
  if DotNet48RebootRequired then
  begin
    NeedsRestart := True;
    Result := '.NET Framework 4.8 安装需要重启计算机才能完成。' + #13#10#13#10 +
      '请重启后重新运行 AccuX 安装程序。';
  end;
  if Result <> '' then
    Log(Result);
end;

procedure RegisterComView(const RegAsm, DllPath, ViewName: String);
var
  ResultCode: Integer;
begin
  Log('注册 COM（' + ViewName + '）：' + DllPath);
  if not Exec(RegAsm, '/codebase "' + DllPath + '"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    FailInstall('无法启动 RegAsm（' + ViewName + '）：' + SysErrorMessage(ResultCode));
  Log(Format('RegAsm（%s）退出码：%d', [ViewName, ResultCode]));
  if ResultCode <> 0 then
    FailInstall(Format('RegAsm（%s）注册失败，退出码 %d；将恢复安装前状态。', [ViewName, ResultCode]));
end;

procedure RegisterComServer;
var
  RegAsm32, RegAsm64, DllPath: String;
begin
  DllPath := ExpandConstant('{app}\AccuX.AddIn.dll');
  RegAsm32 := ExpandConstant('{dotnet4032}\RegAsm.exe');
  if not FileExists(RegAsm32) then
    FailInstall('未找到 32 位 RegAsm.exe：' + RegAsm32);
  if IsWin64 then
  begin
    RegAsm64 := ExpandConstant('{dotnet4064}\RegAsm.exe');
    if not FileExists(RegAsm64) then
      FailInstall('未找到 64 位 RegAsm.exe：' + RegAsm64);
    BackupComView(HKLM32, ExpandConstant('{syswow64}\reg.exe'));
    BackupComView(HKLM64, ExpandConstant('{sys}\reg.exe'));
  end
  else
    BackupComView(HKLM32, ExpandConstant('{sys}\reg.exe'));

  { 必须在启动第一个 RegAsm 前标记：非零退出也可能已写入部分键。
    保存整个专属键树，保留旧版本子键与旧 CodeBase，而不是盲目 /unregister。 }
  ComRegistrationStarted := True;
  RegisterComView(RegAsm32, DllPath, '32 位');
  if IsWin64 then
    RegisterComView(RegAsm64, DllPath, '64 位');
end;

function RegisterComAndGetKey(Param: String): String;
begin
  if not ComRegistrationComplete then
  begin
    RegisterComServer;
    ComRegistrationComplete := True;
  end;
  Result := Param;
end;

procedure DeinitializeSetup;
var
  I, ResultCode: Integer;
  RecoveryFailed: Boolean;
begin
  if InstallCommitted then
    Exit;
  RecoveryFailed := False;
  { Inno 原生撤销完成后再恢复旧文件，避免恢复的文件又被原生撤销删除。 }
  for I := 0 to GetArrayLength(FileBackups) - 1 do
  begin
    if not ForceDirectories(ExtractFileDir(FileBackups[I].Destination)) or
        not CopyFile(FileBackups[I].Backup, FileBackups[I].Destination, False) then
    begin
      Log('恢复旧文件失败：' + FileBackups[I].Destination);
      RecoveryFailed := True;
    end;
  end;
  if ComRegistrationStarted then
    for I := GetArrayLength(RegistryBackups) - 1 downto 0 do
    begin
      if RegKeyExists(RegistryBackups[I].Root, RegistryBackups[I].Subkey) then
        if not RegDeleteKeyIncludingSubkeys(RegistryBackups[I].Root, RegistryBackups[I].Subkey) then
        begin
          Log('清理本次 COM 注册失败：' + RegistryBackups[I].Subkey);
          RecoveryFailed := True;
        end;
      if RegistryBackups[I].Existed then
      begin
        if not Exec(RegistryBackups[I].RegExe, 'import "' + RegistryBackups[I].Backup + '"',
            '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
        begin
          Log('无法启动 COM 注册恢复：' + SysErrorMessage(ResultCode));
          RecoveryFailed := True;
        end
        else if ResultCode <> 0 then
        begin
          Log(Format('恢复 COM 注册失败（%d）：%s', [ResultCode, RegistryBackups[I].Subkey]));
          RecoveryFailed := True;
        end;
      end;
    end;
  if RecoveryFailed then
    ReportInstallError('安装失败后的恢复未完全成功，请根据安装日志修复或重新安装旧版本。')
  else if (GetArrayLength(FileBackups) > 0) or ComRegistrationStarted then
    Log('已恢复旧文件和 COM 注册。');
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then
  begin
    { 在任何文件被覆盖前完成全部备份。保持与 [Files] 清单一致。 }
    BackupPackageFile('AccuX.AddIn.dll');
    BackupPackageFile('AccuX.Core.dll');
    BackupPackageFile('AccuX.Host.dll');
    BackupPackageFile('AccuX.Modules.BasicFinance.dll');
    BackupPackageFile('AccuX.Modules.Mark.dll');
    BackupPackageFile('AccuX.Modules.Compare.dll');
    BackupPackageFile('Newtonsoft.Json.dll');
    BackupPackageFile('Microsoft.Office.Interop.Excel.dll');
    BackupPackageFile('office.dll');
    BackupPackageFile('Microsoft.Vbe.Interop.dll');
    BackupPackageFile('config.sample.json');
  end;
  if CurStep = ssPostInstall then
  begin
    InstallCommitted := True;
    Log('AccuX 文件安装与 COM 注册已提交。');
  end;
end;
