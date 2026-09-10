# 安装器回归前置检查：验证生产脚本的事务入口、文件备份清单和 Inno 编译。
param(
    [string]$Iscc = 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
    [string]$Version = '1.1',
    [string]$FileVersion = '1.1.0.0'
)

$ErrorActionPreference = 'Stop'
$installerRoot = $PSScriptRoot
$repoRoot = Split-Path -Parent $installerRoot
$scriptPath = Join-Path $installerRoot 'AccuX.iss'

if (-not (Test-Path -LiteralPath $Iscc -PathType Leaf)) {
    throw "未找到 Inno Setup 编译器：$Iscc"
}
if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
    throw "未找到安装脚本：$scriptPath"
}

$source = [IO.File]::ReadAllText($scriptPath)
$requiredMarkers = @(
    'RegisterComAndGetKey',
    'RegisterComServer',
    'ComRegistrationStarted := True',
    'procedure DeinitializeSetup',
    'InstallCommitted := True'
)
foreach ($marker in $requiredMarkers) {
    if (-not $source.Contains($marker)) {
        throw "安装脚本缺少事务安全入口：$marker"
    }
}

$unsupportedSilentMarkers = @(
    'WizardSilent',
    'SuppressibleMsgBox',
    '/SILENT',
    '/VERYSILENT'
)
foreach ($marker in $unsupportedSilentMarkers) {
    if ($source.Contains($marker)) {
        throw "安装脚本不应包含静默安装分支：$marker"
    }
}

$sourceFiles = [regex]::Matches($source, 'Source:\s+"([^"]+)";') |
    ForEach-Object { $_.Groups[1].Value }
$backupFiles = [regex]::Matches($source, "BackupPackageFile\('([^']+)'\)") |
    ForEach-Object { $_.Groups[1].Value }

foreach ($sourceFile in $sourceFiles) {
    $relative = $sourceFile.Replace('{#SourceRoot}', '..\src')
    $resolved = [IO.Path]::GetFullPath((Join-Path $installerRoot $relative))
    if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
        throw "安装包输入文件不存在：$resolved"
    }

    $fileName = [IO.Path]::GetFileName($relative)
    if (-not ($backupFiles | Where-Object { $_ -eq $fileName })) {
        throw "文件备份清单遗漏：$fileName"
    }
}

$testId = [Guid]::NewGuid().ToString('N')
$testRoot = Join-Path $repoRoot "artifacts\installer-tests\$testId"
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null

try {
    $versionKeys = @(
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Inno Setup 6_is1',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Inno Setup 6_is1'
    )
    $isccVersion = $null
    foreach ($versionKey in $versionKeys) {
        if (Test-Path -LiteralPath $versionKey) {
            $isccVersion = (Get-ItemProperty -LiteralPath $versionKey).DisplayVersion
            break
        }
    }
    if ($isccVersion -ne '6.7.3') {
        throw "安装器回归测试要求 Inno Setup 6.7.3，实际为 $isccVersion。"
    }

    $validationOutput = Join-Path $testRoot 'AccuXSetup-validation'
    & $Iscc `
        "/DAccuXVersion=$Version" `
        "/DAccuXFileVersion=$FileVersion" `
        '/DAccuXOutputBaseFilename=AccuXSetup-validation' `
        "/O$testRoot" `
        $scriptPath
    if ($LASTEXITCODE -ne 0) {
        throw 'Inno Setup 安装脚本编译失败。'
    }
    $validationExe = "$validationOutput.exe"
    if (-not (Test-Path -LiteralPath $validationExe -PathType Leaf)) {
        throw "未生成验证安装包：$validationExe"
    }

    $report = @(
        "InnoSetupVersion=$isccVersion"
        "Version=$Version"
        "SourceFileCount=$($sourceFiles.Count)"
        "BackupFileCount=$($backupFiles.Count)"
        "ValidationInstaller=$validationExe"
        'TransactionMarkers=passed'
    )
    $report | Set-Content -LiteralPath (Join-Path $testRoot 'validation.txt') -Encoding utf8
    Write-Host "安装器回归前置检查通过：$testRoot" -ForegroundColor Green
}
catch {
    $_ | Out-String | Set-Content -LiteralPath (Join-Path $testRoot 'failure.log') -Encoding utf8
    throw
}
