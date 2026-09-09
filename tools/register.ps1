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

$regAsm32 = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\RegAsm.exe'
$regAsm64 = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
$regExe = Join-Path $env:WINDIR 'System32\reg.exe'

$registryViews = @(
    @{ Name = '32'; RegAsm = $regAsm32 }
)
if (Test-Path $regAsm64) {
    $registryViews += @{ Name = '64'; RegAsm = $regAsm64 }
}
foreach ($view in $registryViews) {
    if (-not (Test-Path $view.RegAsm)) {
        throw "找不到 $($view.Name) 位 RegAsm.exe（需要 .NET Framework 4.x）。"
    }
}

foreach ($view in $registryViews) {
    $tempReg = Join-Path $env:TEMP ("accux-$($view.Name)-" + [Guid]::NewGuid().ToString('N') + ".reg")
    try {
        Write-Host "生成 $($view.Name) 位 COM 注册表脚本 ..."
        & $view.RegAsm /regfile:$tempReg /codebase $addInDll | Out-Null
        if (-not (Test-Path $tempReg)) {
            throw "RegAsm 未能生成 $($view.Name) 位注册表脚本。"
        }

        # RegAsm 生成的是 ANSI 编码的 REGEDIT4 脚本；reg import 需要 v5.00 格式的 Unicode 文件。
        # 同时把机器级 HKCR 改写为当前用户 HKCU\Software\Classes，实现按用户注册。
        $content = Get-Content -Raw -Encoding Default $tempReg
        $content = $content -replace 'REGEDIT4', 'Windows Registry Editor Version 5.00'
        $content = $content -replace 'HKEY_CLASSES_ROOT', 'HKEY_CURRENT_USER\Software\Classes'
        [System.IO.File]::WriteAllText($tempReg, $content, [System.Text.Encoding]::Unicode)

        Write-Host "导入 $($view.Name) 位 COM 注册项（HKCU）..."
        & $regExe import $tempReg "/reg:$($view.Name)" | Out-Null
        if ($LASTEXITCODE -ne 0) {
            throw "$($view.Name) 位 reg import 失败（退出码 $LASTEXITCODE）。"
        }
    }
    finally {
        if (Test-Path $tempReg) {
            Remove-Item $tempReg -Force -ErrorAction SilentlyContinue
        }
    }
}

# Excel 使用 ProgId 子键；WPS 表格使用 AddinsWL 根键下的 ProgId 值。
foreach ($view in $registryViews) {
    & $regExe add "HKCU\Software\Microsoft\Office\Excel\Addins\$progId" /v Description /t REG_SZ /d 'AccuX 财务效率插件' /f "/reg:$($view.Name)" | Out-Null
    & $regExe add "HKCU\Software\Microsoft\Office\Excel\Addins\$progId" /v FriendlyName /t REG_SZ /d $friendlyName /f "/reg:$($view.Name)" | Out-Null
    & $regExe add "HKCU\Software\Microsoft\Office\Excel\Addins\$progId" /v LoadBehavior /t REG_DWORD /d 3 /f "/reg:$($view.Name)" | Out-Null
}

$wpsRoot = 'HKCU:\Software\Kingsoft\Office\ET\AddinsWL'
if (-not (Test-Path $wpsRoot)) {
    New-Item -Path $wpsRoot -Force | Out-Null
}
New-ItemProperty -Path $wpsRoot -Name $progId -Value '' -PropertyType String -Force | Out-Null

Write-Host "注册完成。ProgId = $progId" -ForegroundColor Green
Write-Host '请完全退出并重新打开 WPS 表格，再查看“开发工具 → COM 加载项”。'
