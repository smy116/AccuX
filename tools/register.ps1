# AccuX 开发期注册脚本（当前用户 HKCU，无需管理员）
#
# 用法：
#   powershell -ExecutionPolicy Bypass -File tools\register.ps1
#   powershell -ExecutionPolicy Bypass -File tools\register.ps1 -Configuration Release
#
# 实现说明：
#   - RegAsm /regfile 只生成注册表脚本（不写系统），再把根键 HKEY_CLASSES_ROOT
#     改写为 HKEY_CURRENT_USER\Software\Classes 后导入，实现真正的按用户注册、免管理员。
#   - Excel / WPS 的 Addins 子键名必须使用 ProgId（宿主按子键名做 CoCreateInstance）。
#   - 开发期与安装包必须使用同一套 ProgId / LoadBehavior，避免“VS 能调试、安装包无法加载”。
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug'
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$addInDll = Join-Path $repoRoot "src\AccuX.AddIn\bin\$Configuration\net48\AccuX.AddIn.dll"

if (-not (Test-Path $addInDll)) {
    throw "找不到 Add-in 程序集：$addInDll。请先构建 $Configuration 配置。"
}

$progId = 'AccuX.AddIn.Connect'
$friendlyName = 'AccuX'

$regAsm = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
if (-not (Test-Path $regAsm)) {
    $regAsm = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\RegAsm.exe'
}
if (-not (Test-Path $regAsm)) {
    throw '找不到 RegAsm.exe（需要 .NET Framework 4.x）。'
}

$tempReg = Join-Path $env:TEMP ("accux-" + [Guid]::NewGuid().ToString('N') + ".reg")

try {
    Write-Host "生成注册表脚本 ..."
    & $regAsm /regfile:$tempReg /codebase $addInDll | Out-Null
    if (-not (Test-Path $tempReg)) {
        throw 'RegAsm 未能生成注册表脚本。'
    }

    # RegAsm 生成的是 ANSI 编码的 REGEDIT4 脚本；reg import 需要 v5.00 格式的 Unicode 文件。
    # 同时把机器级 HKCR 改写为当前用户 HKCU\Software\Classes，实现按用户注册。
    $content = Get-Content -Raw -Encoding Default $tempReg
    $content = $content -replace 'REGEDIT4', 'Windows Registry Editor Version 5.00'
    $content = $content -replace 'HKEY_CLASSES_ROOT', 'HKEY_CURRENT_USER\Software\Classes'
    [System.IO.File]::WriteAllText($tempReg, $content, [System.Text.Encoding]::Unicode)

    Write-Host "导入 COM 注册项（HKCU）..."
    & reg import $tempReg | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "reg import 失败（退出码 $LASTEXITCODE）。"
    }

    # Excel 与 WPS 表格共用同一 ProgId；子键名必须是 ProgId。
    $addInRoots = @(
        'HKCU:\Software\Microsoft\Office\Excel\Addins',
        'HKCU:\Software\Kingsoft\Office\ET\AddinsWL'
    )

    foreach ($root in $addInRoots) {
        $key = Join-Path $root $progId
        if (-not (Test-Path $key)) {
            New-Item -Path $key -Force | Out-Null
        }

        New-ItemProperty -Path $key -Name 'Description'  -Value 'AccuX 财务效率插件' -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $key -Name 'FriendlyName' -Value $friendlyName        -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $key -Name 'LoadBehavior' -Value 3                    -PropertyType DWord  -Force | Out-Null
    }

    Write-Host "注册完成。ProgId = $progId" -ForegroundColor Green
    Write-Host '现在可以在 Visual Studio 中按 F5 启动 Excel 调试，或直接打开 Excel 查看 AccuX 选项卡。'
}
finally {
    if (Test-Path $tempReg) {
        Remove-Item $tempReg -Force -ErrorAction SilentlyContinue
    }
}
