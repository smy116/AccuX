# AccuX 开发期注销脚本（当前用户 HKCU，无需管理员）
#
# 用法：
#   powershell -ExecutionPolicy Bypass -File tools\unregister.ps1
#   powershell -ExecutionPolicy Bypass -File tools\unregister.ps1 -Configuration Release
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug'
)

$ErrorActionPreference = 'Stop'

$progId = 'AccuX.AddIn.Connect'

# 1. 删除 Add-in 加载项注册项（Excel / WPS）。
$addInRoots = @(
    'HKCU:\Software\Microsoft\Office\Excel\Addins',
    'HKCU:\Software\Kingsoft\Office\ET\AddinsWL'
)

foreach ($root in $addInRoots) {
    $key = Join-Path $root $progId
    if (Test-Path $key) {
        Remove-Item -Path $key -Recurse -Force
        Write-Host "已删除 $key"
    }
}

# 2. 删除当前用户下的 COM 注册项。
$classesRoot = 'HKCU:\Software\Classes'
$classKeys = @(
    (Join-Path $classesRoot $progId),
    (Join-Path $classesRoot ('CLSID\{7C1F0E4A-9B2D-4E8C-A1F3-6D5E8B0C2A11}'))
)

foreach ($key in $classKeys) {
    if (Test-Path $key) {
        Remove-Item -Path $key -Recurse -Force
        Write-Host "已删除 $key"
    }
}

Write-Host '注销完成。' -ForegroundColor Green
