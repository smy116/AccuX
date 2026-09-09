; AccuX V1 安装脚本（规格 §26 安装与部署约定）
;
; 编译：ISCC.exe installer\AccuX.iss
; 产物：installer\Output\AccuXSetup-1.0.0.exe
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

#define AccuXVersion "1.0.0"
#define AccuXProgId "AccuX.AddIn.Connect"
#define AccuXFriendlyName "AccuX"
#define SourceRoot "..\src"

; Add-ins 子键名必须是 ProgId：宿主按子键名做 CoCreateInstance。
#define ExcelAddinKey "Software\Microsoft\Office\Excel\Addins\" + AccuXProgId
#define WpsAddinKey "Software\Kingsoft\Office\ET\AddinsWL\" + AccuXProgId

[Setup]
AppId={{8E3B2A64-1C7D-4A9F-9E5B-2D6F0A8C4B31}
AppName=AccuX
AppVersion={#AccuXVersion}
AppVerName=AccuX {#AccuXVersion}
AppPublisher=AccuX
DefaultDirName={autopf}\AccuX
DefaultGroupName=AccuX
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=AccuXSetup-{#AccuXVersion}
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=admin
MinVersion=6.1sp1
; 64 位 Windows 上使用原生 64 位 Program Files 与 64 位注册表视图；
; 32 位 Windows 上自动回退为 32 位安装模式（不设 ArchitecturesAllowed）。
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
UninstallDisplayName=AccuX {#AccuXVersion}

[Languages]
; 先读官方默认消息，再由简体中文覆盖，保证个别未翻译消息仍有英文兜底。
Name: "chinese"; MessagesFile: "compiler:Default.isl,Languages\ChineseSimplified.isl"

[Files]
; AccuX 程序集与随附的 Office PIA（Private 部署，不依赖目标机 GAC 布局）
Source: "{#SourceRoot}\AccuX.AddIn\bin\Release\net48\AccuX.AddIn.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\AccuX.AddIn\bin\Release\net48\AccuX.Core.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\AccuX.AddIn\bin\Release\net48\AccuX.Host.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\AccuX.AddIn\bin\Release\net48\AccuX.Modules.BasicFinance.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\AccuX.AddIn\bin\Release\net48\Newtonsoft.Json.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\AccuX.AddIn\bin\Release\net48\Microsoft.Office.Interop.Excel.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\AccuX.AddIn\bin\Release\net48\office.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\AccuX.AddIn\bin\Release\net48\Microsoft.Vbe.Interop.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceRoot}\AccuX.AddIn\config.sample.json"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; Excel Add-in 注册（机器级，安装包需要管理员权限）。
; 同一组值分别写入默认视图（64 位安装模式下即 64 位视图）与 32 位视图，
; 保证 64 位 Office 与 32 位 Office 都能读到。
Root: HKLM;   Subkey: "{#ExcelAddinKey}"; ValueType: string; ValueName: "Description";  ValueData: "AccuX 财务效率插件"; Flags: uninsdeletekey
Root: HKLM;   Subkey: "{#ExcelAddinKey}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AccuXFriendlyName}"
Root: HKLM;   Subkey: "{#ExcelAddinKey}"; ValueType: dword;  ValueName: "LoadBehavior"; ValueData: "3"
Root: HKLM32; Subkey: "{#ExcelAddinKey}"; ValueType: string; ValueName: "Description";  ValueData: "AccuX 财务效率插件"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "{#ExcelAddinKey}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AccuXFriendlyName}"
Root: HKLM32; Subkey: "{#ExcelAddinKey}"; ValueType: dword;  ValueName: "LoadBehavior"; ValueData: "3"

; WPS 表格 Add-in 注册（同样写入两个注册表视图）。
Root: HKLM;   Subkey: "{#WpsAddinKey}"; ValueType: string; ValueName: "Description";  ValueData: "AccuX 财务效率插件"; Flags: uninsdeletekey
Root: HKLM;   Subkey: "{#WpsAddinKey}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AccuXFriendlyName}"
Root: HKLM;   Subkey: "{#WpsAddinKey}"; ValueType: dword;  ValueName: "LoadBehavior"; ValueData: "3"
Root: HKLM32; Subkey: "{#WpsAddinKey}"; ValueType: string; ValueName: "Description";  ValueData: "AccuX 财务效率插件"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "{#WpsAddinKey}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AccuXFriendlyName}"
Root: HKLM32; Subkey: "{#WpsAddinKey}"; ValueType: dword;  ValueName: "LoadBehavior"; ValueData: "3"

[UninstallRun]
; 注销 COM 可见程序集（与 [Code] 中注册逻辑对应）。
Filename: "{dotnet4032}\RegAsm.exe"; Parameters: "/unregister ""{app}\AccuX.AddIn.dll"""; Flags: runhidden waituntilterminated; RunOnceId: "UnregAsm32"
Filename: "{dotnet4064}\RegAsm.exe"; Parameters: "/unregister ""{app}\AccuX.AddIn.dll"""; Flags: runhidden waituntilterminated; RunOnceId: "UnregAsm64"; Check: IsWin64

[Code]
const
  { .NET Framework 4.8 的 Release 值：
    4.8 = 528040（Win10 1903/1909）、528049（其他系统）、528372（Win10 2004/20H2）；
    4.8.1 = 533325 起。按微软建议使用 >= 比较。 }
  DotNet48MinRelease = 528040;

  { .NET Framework 4.8 Web 安装程序（ndp48-web.exe，约 1.5 MB）。
    fwlink 会 302 重定向到 download.visualstudio.microsoft.com，Inno 下载函数自动跟随。
    若微软重新发布该引导程序导致 SHA-256 变化，下载会失败并给出人工安装提示，
    此时请更新下面的 Ndp48Sha256 常量。 }
  Ndp48Url = 'https://go.microsoft.com/fwlink/?LinkId=2085155';
  Ndp48BaseName = 'ndp48-web.exe';
  Ndp48Sha256 = '0bba3094588c4bfec301939985222a20b340bf03431563dec8b2b4478b06fffa';

var
  DownloadPage: TDownloadWizardPage;
  DotNet48RebootRequired: Boolean;

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

{ 下载并静默安装 .NET Framework 4.8。返回 True 表示已就绪或已交由重启完成。 }
function InstallDotNet48(): Boolean;
var
  InstallerPath: String;
  ResultCode: Integer;
  ErrorMessage: String;
begin
  Result := False;

  DownloadPage.Clear;
  DownloadPage.Add(Ndp48Url, Ndp48BaseName, Ndp48Sha256);
  DownloadPage.Show;
  try
    try
      DownloadPage.Download;
    except
      if DownloadPage.AbortedByUser then
        MsgBox('已取消下载 .NET Framework 4.8 安装程序。', mbError, MB_OK)
      else
      begin
        ErrorMessage := Format('下载 .NET Framework 4.8 安装程序失败：%s', [GetExceptionMessage]);
        ErrorMessage := ErrorMessage + #13#10#13#10 +
          '请检查网络连接后重试，或从以下地址手动安装后重新运行本安装程序：' + #13#10 +
          'https://dotnet.microsoft.com/download/dotnet-framework/net48';
        MsgBox(ErrorMessage, mbCriticalError, MB_OK);
      end;
      Exit;
    end;
  finally
    DownloadPage.Hide;
  end;

  InstallerPath := ExpandConstant('{tmp}\' + Ndp48BaseName);
  if not FileExists(InstallerPath) then
  begin
    MsgBox('未找到已下载的 .NET Framework 4.8 安装程序。', mbCriticalError, MB_OK);
    Exit;
  end;

  if not Exec(InstallerPath, '/q /norestart /ChainingPackage AccuX', '',
      SW_SHOW, ewWaitUntilTerminated, ResultCode) then
  begin
    MsgBox('无法启动 .NET Framework 4.8 安装程序。', mbCriticalError, MB_OK);
    Exit;
  end;

  { 0 = 成功；3010 / 1641 = 成功但需重启；1602 = 用户取消；1603 = 致命错误；
    5100 = 不满足系统要求。 }
  case ResultCode of
    0, 3010, 1641:
      begin
        if IsDotNet48OrLater() then
          Result := True
        else
        begin
          { 安装程序返回成功但注册表尚未反映，通常需要重启后才能完成。 }
          DotNet48RebootRequired := True;
          Result := True;
        end;
      end;
    1602:
      MsgBox('.NET Framework 4.8 安装已被取消。', mbError, MB_OK);
    1603:
      MsgBox('.NET Framework 4.8 安装失败（错误 1603：安装过程中发生致命错误）。' + #13#10#13#10 +
        '请查看安装日志，或手动安装后重新运行本安装程序。', mbCriticalError, MB_OK);
    5100:
      MsgBox('.NET Framework 4.8 安装失败（错误 5100：本机不满足系统要求）。', mbCriticalError, MB_OK);
  else
    MsgBox(Format('.NET Framework 4.8 安装程序返回错误代码 %d。', [ResultCode]),
      mbCriticalError, MB_OK);
  end;
end;

{ 缺少 .NET Framework 4.8 时在「准备安装」页之前完成下载安装。 }
function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if CurPageID <> wpReady then
    Exit;

  if IsDotNet48OrLater() then
    Exit;

  if MsgBox('AccuX 需要 .NET Framework 4.8 或更高版本，当前系统未检测到。' + #13#10#13#10 +
      '是否立即联网下载并安装？（约 1.5 MB，需要网络连接）',
      mbConfirmation, MB_YESNO) = IDNO then
  begin
    MsgBox('未安装 .NET Framework 4.8，AccuX 安装已中止。', mbError, MB_OK);
    Result := False;
    Exit;
  end;

  if not InstallDotNet48() then
  begin
    MsgBox('未能完成 .NET Framework 4.8 安装，AccuX 安装已中止。', mbCriticalError, MB_OK);
    Result := False;
  end;
end;

{ .NET Framework 安装要求重启时，先让用户重启再重新运行，避免注册到不完整运行时。 }
function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  if DotNet48RebootRequired then
  begin
    NeedsRestart := True;
    Result := '.NET Framework 4.8 安装需要重启计算机才能完成。' + #13#10#13#10 +
      '请重启后重新运行 AccuX 安装程序。';
  end;
end;

{ 使用 RegAsm 注册 COM 可见程序集（与 tools\register.ps1 使用同一 ProgId）。
  32 位与 64 位分别注册，覆盖两个注册表视图。注册失败即中止安装并回滚。 }
procedure RegisterComServer();
var
  DllPath: String;
  RegAsm: String;
  ResultCode: Integer;
begin
  DllPath := ExpandConstant('{app}\AccuX.AddIn.dll');

  RegAsm := ExpandConstant('{dotnet4032}\RegAsm.exe');
  if not FileExists(RegAsm) then
    RaiseException('未找到 32 位 RegAsm.exe：' + RegAsm);
  if not Exec(RegAsm, '/codebase "' + DllPath + '"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    RaiseException('无法启动 32 位 RegAsm.exe：' + RegAsm);
  if ResultCode <> 0 then
    RaiseException(Format('RegAsm（32 位）注册失败，退出码 %d。', [ResultCode]));

  if IsWin64 then
  begin
    RegAsm := ExpandConstant('{dotnet4064}\RegAsm.exe');
    if not FileExists(RegAsm) then
      RaiseException('未找到 64 位 RegAsm.exe：' + RegAsm);
    if not Exec(RegAsm, '/codebase "' + DllPath + '"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
      RaiseException('无法启动 64 位 RegAsm.exe：' + RegAsm);
    if ResultCode <> 0 then
      RaiseException(Format('RegAsm（64 位）注册失败，退出码 %d。', [ResultCode]));
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    RegisterComServer();
end;
