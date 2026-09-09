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

$clsid = '{7C1F0E4A-9B2D-4E8C-A1F3-6D5E8B0C2A11}'
$regAsm64 = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
$registryViews = @('32')
if (Test-Path $regAsm64) {
    $registryViews += '64'
}

function Remove-RegistryKeyForView {
    param(
        [string]$ViewName,
        [string]$SubKey
    )

    $view = if ($ViewName -eq '32') {
        [Microsoft.Win32.RegistryView]::Registry32
    }
    else {
        [Microsoft.Win32.RegistryView]::Registry64
    }

    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::CurrentUser,
        $view
    )
    try {
        $existing = $base.OpenSubKey($SubKey)
        if ($null -ne $existing) {
            $existing.Dispose()
            $base.DeleteSubKeyTree($SubKey)
            Write-Host "已删除 $ViewName 位 HKCU\$SubKey"
        }
    }
    finally {
        $base.Dispose()
    }
}

function Remove-RegistryValueForView {
    param(
        [string]$ViewName,
        [string]$SubKey,
        [string]$ValueName
    )

    $view = if ($ViewName -eq '32') {
        [Microsoft.Win32.RegistryView]::Registry32
    }
    else {
        [Microsoft.Win32.RegistryView]::Registry64
    }

    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::CurrentUser,
        $view
    )
    try {
        $key = $base.OpenSubKey($SubKey, $true)
        if ($null -ne $key) {
            try {
                if ($key.GetValueNames() -contains $ValueName) {
                    $key.DeleteValue($ValueName)
                    Write-Host "已删除 $ViewName 位 HKCU\$SubKey\$ValueName"
                }
            }
            finally {
                $key.Dispose()
            }
        }
    }
    finally {
        $base.Dispose()
    }
}

foreach ($viewName in $registryViews) {
    # Excel 使用 ProgId 子键；同时清理旧版本误写的 WPS 子键。
    Remove-RegistryKeyForView $viewName "Software\Microsoft\Office\Excel\Addins\$progId"
    Remove-RegistryKeyForView $viewName "Software\Kingsoft\Office\ET\AddinsWL\$progId"

    # WPS 使用 AddinsWL 根键下的 ProgId 值。
    Remove-RegistryValueForView $viewName 'Software\Kingsoft\Office\ET\AddinsWL' $progId

    # RegAsm /regfile 导入到 HKCU\Software\Classes 的 COM 注册项。
    Remove-RegistryKeyForView $viewName "Software\Classes\$progId"
    Remove-RegistryKeyForView $viewName "Software\Classes\CLSID\$clsid"
}

Write-Host '注销完成。' -ForegroundColor Green
