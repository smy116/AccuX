# 安装器回归检查：验证生产脚本的事务入口、文件备份清单，以及正式/预发布两种版本的 PE 元数据。
param(
    [string]$Iscc = 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
    [string]$Version = '1.6.1',
    [string]$FileVersion = '1.6.1.0'
)

$ErrorActionPreference = 'Stop'
$installerRoot = $PSScriptRoot
$repoRoot = Split-Path -Parent $installerRoot
$scriptPath = Join-Path $installerRoot 'AccuX.iss'
$builderPath = Join-Path $installerRoot 'Build-Installer.ps1'

if (-not (Test-Path -LiteralPath $Iscc -PathType Leaf)) {
    throw "未找到 Inno Setup 编译器：$Iscc"
}
if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
    throw "未找到安装脚本：$scriptPath"
}
if (-not (Test-Path -LiteralPath $builderPath -PathType Leaf)) {
    throw "未找到统一安装包构建脚本：$builderPath"
}

$source = [IO.File]::ReadAllText($scriptPath)
$requiredMarkers = @(
    'RegisterComAndGetKey',
    'RegisterComServer',
    'ComRegistrationStarted := True',
    'procedure DeinitializeSetup',
    'InstallCommitted := True',
    '#ifndef AccuXDisplayVersion',
    '#ifndef AccuXInstallerVersion',
    '#ifndef AccuXFileVersion',
    'AppVersion={#AccuXInstallerVersion}',
    'AppVerName=AccuX {#AccuXDisplayVersion}',
    'VersionInfoVersion={#AccuXFileVersion}',
    'VersionInfoProductVersion={#AccuXFileVersion}',
    'VersionInfoTextVersion={#AccuXDisplayVersion}',
    'VersionInfoProductTextVersion={#AccuXDisplayVersion}'
)
foreach ($marker in $requiredMarkers) {
    if (-not $source.Contains($marker)) {
        throw "安装脚本缺少事务或版本元数据入口：$marker"
    }
}
if ($source.Contains('AppVersion={#AccuXDisplayVersion}') -or
    $source.Contains('VersionInfoVersion={#AccuXDisplayVersion}') -or
    $source.Contains('VersionInfoProductVersion={#AccuXDisplayVersion}')) {
    throw '预发布显示版本不能进入数字版本字段。'
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

function ConvertTo-FourPartVersion {
    param([Parameter(Mandatory = $true)][string]$Value)

    $parts = $Value.Split('-')[0].Split('.')
    $patch = if ($parts.Length -ge 3) { $parts[2] } else { '0' }
    return "$($parts[0]).$($parts[1]).$patch.0"
}

function Get-NumericVersion {
    param([Parameter(Mandatory = $true)][string]$Value)

    return $Value.Split('-')[0]
}

$numericVersion = Get-NumericVersion -Value $Version
$scenarios = @(
    [pscustomobject]@{
        Label = 'requested'
        DisplayVersion = $Version
        FileVersion = $FileVersion
    }
)
if ($Version.Contains('-')) {
    $scenarios += [pscustomobject]@{
        Label = 'stable'
        DisplayVersion = $numericVersion
        FileVersion = ConvertTo-FourPartVersion -Value $numericVersion
    }
}
else {
    $scenarios += [pscustomobject]@{
        Label = 'ci'
        DisplayVersion = "$numericVersion-ci.1.test"
        FileVersion = (ConvertTo-FourPartVersion -Value $numericVersion) -replace '\.0$', '.1'
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

    $reports = @()
    foreach ($scenario in $scenarios) {
        $outputDirectory = Join-Path $testRoot $scenario.Label
        $outputBase = "AccuXSetup-validation-$($scenario.Label)"
        & $builderPath `
            -Iscc $Iscc `
            -Version $scenario.DisplayVersion `
            -FileVersion $scenario.FileVersion `
            -OutputDirectory $outputDirectory `
            -OutputBaseFilename $outputBase `
            -CommitSha 'installer-test' `
            -RunNumber $testId
        if ($LASTEXITCODE -ne 0) {
            throw "Inno Setup 安装脚本编译失败：$($scenario.Label)"
        }

        $validationExe = Join-Path $outputDirectory "$outputBase.exe"
        $checksumPath = "$validationExe.sha256"
        $manifestPath = "$validationExe.manifest.json"
        foreach ($path in @($validationExe, $checksumPath, $manifestPath)) {
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                throw "未生成验证产物：$path"
            }
        }

        $bytes = [IO.File]::ReadAllBytes($validationExe)
        if ($bytes.Length -lt 64 -or $bytes[0] -ne 0x4d -or $bytes[1] -ne 0x5a) {
            throw "验证安装包不是有效的 MZ 文件：$validationExe"
        }
        $peOffset = [BitConverter]::ToInt32($bytes, 0x3c)
        if ($peOffset -lt 0 -or $peOffset + 4 -gt $bytes.Length -or
            $bytes[$peOffset] -ne 0x50 -or $bytes[$peOffset + 1] -ne 0x45 -or
            $bytes[$peOffset + 2] -ne 0 -or $bytes[$peOffset + 3] -ne 0) {
            throw "验证安装包不是有效的 PE 文件：$validationExe"
        }
        $info = [Diagnostics.FileVersionInfo]::GetVersionInfo($validationExe)
        $fixedFileVersion = "$($info.FileMajorPart).$($info.FileMinorPart).$($info.FileBuildPart).$($info.FilePrivatePart)"
        $fixedProductVersion = "$($info.ProductMajorPart).$($info.ProductMinorPart).$($info.ProductBuildPart).$($info.ProductPrivatePart)"
        if ($fixedFileVersion -ne $scenario.FileVersion -or $fixedProductVersion -ne $scenario.FileVersion) {
            throw "固定版本字段不一致：$($scenario.Label) File=$fixedFileVersion Product=$fixedProductVersion"
        }
        if ($info.FileVersion.Trim() -ne $scenario.DisplayVersion -or $info.ProductVersion.Trim() -ne $scenario.DisplayVersion) {
            throw "文本版本字段不一致：$($scenario.Label)"
        }

        $hash = (Get-FileHash -LiteralPath $validationExe -Algorithm SHA256).Hash.ToLowerInvariant()
        $checksumLine = [IO.File]::ReadAllText($checksumPath).Trim()
        if ($checksumLine -ne "$hash *$([IO.Path]::GetFileName($validationExe))") {
            throw "校验文件与验证安装包不一致：$($scenario.Label)"
        }
        $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
        $expectedInstallerVersion = Get-NumericVersion -Value $scenario.DisplayVersion
        foreach ($pair in @{
            version = $scenario.DisplayVersion
            displayVersion = $scenario.DisplayVersion
            installerVersion = $expectedInstallerVersion
            fileVersion = $scenario.FileVersion
            fileName = [IO.Path]::GetFileName($validationExe)
            sha256 = $hash
            runNumber = $testId
        }.GetEnumerator()) {
            if ([string]$manifest.($pair.Key) -ne [string]$pair.Value) {
                throw "manifest 字段 $($pair.Key) 不一致：$($scenario.Label)"
            }
        }

        $reports += "Scenario=$($scenario.Label);DisplayVersion=$($scenario.DisplayVersion);InstallerVersion=$expectedInstallerVersion;FileVersion=$($scenario.FileVersion);Hash=$hash;ValidationInstaller=$validationExe"
    }

    $report = @(
        "InnoSetupVersion=$isccVersion"
        "SourceFileCount=$($sourceFiles.Count)"
        "BackupFileCount=$($backupFiles.Count)"
        'VersionFieldMapping=passed'
        'PEAndHashValidation=passed'
        'TransactionMarkers=passed'
    ) + $reports
    $report | Set-Content -LiteralPath (Join-Path $testRoot 'validation.txt') -Encoding utf8
    Write-Host "安装器回归检查通过（正式/预发布）：$testRoot" -ForegroundColor Green
}
catch {
    $_ | Out-String | Set-Content -LiteralPath (Join-Path $testRoot 'failure.log') -Encoding utf8
    throw
}
