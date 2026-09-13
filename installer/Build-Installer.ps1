<#
    AccuX 安装包的唯一构建入口。

    该脚本同时供本地构建和 GitHub Actions 使用。安装器的显示版本、
    Inno Setup 的数字版本和 PE 文件版本必须由这里统一传给 AccuX.iss，
    以避免预发布 SemVer 被写入 Windows 只接受数字的版本字段。
#>
[CmdletBinding()]
param(
    [string]$Iscc = 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
    [Parameter(Mandatory = $true)]
    [string]$Version,
    [Parameter(Mandatory = $true)]
    [string]$FileVersion,
    [string]$OutputDirectory = '',
    [string]$OutputBaseFilename = '',
    [string]$InstallerVersion = '',
    [string]$CommitSha = '',
    [string]$RunNumber = ''
)

$ErrorActionPreference = 'Stop'
$installerRoot = $PSScriptRoot
$scriptPath = Join-Path $installerRoot 'AccuX.iss'

function Assert-VersionSegment {
    param([Parameter(Mandatory = $true)][string]$Value)

    foreach ($part in $Value.Split('.')) {
        if ([int64]$part -gt 65534) {
            throw "版本段必须在 0 到 65534 之间：$Value"
        }
    }
}

function Assert-SemanticVersion {
    param([Parameter(Mandatory = $true)][string]$Value)

    if ($Value -notmatch '^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(?:\.(0|[1-9][0-9]*))?(?:-[0-9A-Za-z][0-9A-Za-z.\-]*)?$') {
        throw "Version 必须是两段或三段数字，可带 -预发布后缀；实际值：$Value"
    }

    Assert-VersionSegment -Value $Value.Split('-')[0]
}

function Assert-FourPartVersion {
    param(
        [Parameter(Mandatory = $true)][string]$Value,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($Value -notmatch '^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)$') {
        throw "$Name 必须是四段数字且不能有前导零；实际值：$Value"
    }

    Assert-VersionSegment -Value $Value
}

function Get-NumericInstallerVersion {
    param([Parameter(Mandatory = $true)][string]$Value)

    $numericParts = $Value.Split('-')[0].Split('.')
    $numeric = if ($numericParts.Length -eq 2) {
        "$($numericParts[0]).$($numericParts[1])"
    }
    else {
        "$($numericParts[0]).$($numericParts[1]).$($numericParts[2])"
    }
    Assert-VersionSegment -Value $numeric
    return $numeric
}

function Assert-OutputBaseFilename {
    param([Parameter(Mandatory = $true)][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value) -or $Value -match '[\\/:*?"<>|]' -or $Value -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
        throw "OutputBaseFilename 不是有效的文件名基名：$Value"
    }
}

function Assert-InstallerBinary {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ExpectedFileVersion,
        [Parameter(Mandatory = $true)][string]$ExpectedDisplayVersion
    )

    $bytes = [IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -lt 64 -or $bytes[0] -ne 0x4d -or $bytes[1] -ne 0x5a) {
        throw "安装包不是有效的 MZ/PE 文件：$Path"
    }

    $peOffset = [BitConverter]::ToInt32($bytes, 0x3c)
    if ($peOffset -lt 0 -or $peOffset + 4 -gt $bytes.Length -or
        $bytes[$peOffset] -ne 0x50 -or $bytes[$peOffset + 1] -ne 0x45 -or
        $bytes[$peOffset + 2] -ne 0 -or $bytes[$peOffset + 3] -ne 0) {
        throw "安装包缺少有效的 PE 头：$Path"
    }

    $info = [Diagnostics.FileVersionInfo]::GetVersionInfo($Path)
    $fixedFileVersion = "$($info.FileMajorPart).$($info.FileMinorPart).$($info.FileBuildPart).$($info.FilePrivatePart)"
    $fixedProductVersion = "$($info.ProductMajorPart).$($info.ProductMinorPart).$($info.ProductBuildPart).$($info.ProductPrivatePart)"
    if ($fixedFileVersion -ne $ExpectedFileVersion) {
        throw "安装包固定 FileVersion 不一致：期望 $ExpectedFileVersion，实际 $fixedFileVersion。"
    }
    if ($fixedProductVersion -ne $ExpectedFileVersion) {
        throw "安装包固定 ProductVersion 不一致：期望 $ExpectedFileVersion，实际 $fixedProductVersion。"
    }

    # VersionInfoTextVersion / VersionInfoProductTextVersion 是面向用户的字符串；
    # 预发布后缀只能出现在这两个字段中，不能污染固定数字版本。
    if ($info.FileVersion.Trim() -ne $ExpectedDisplayVersion) {
        throw "安装包文本 FileVersion 不一致：期望 $ExpectedDisplayVersion，实际 $($info.FileVersion)。"
    }
    if ($info.ProductVersion.Trim() -ne $ExpectedDisplayVersion) {
        throw "安装包文本 ProductVersion 不一致：期望 $ExpectedDisplayVersion，实际 $($info.ProductVersion)。"
    }

    return [pscustomobject]@{
        FixedFileVersion = $fixedFileVersion
        FixedProductVersion = $fixedProductVersion
        TextFileVersion = $info.FileVersion.Trim()
        TextProductVersion = $info.ProductVersion.Trim()
    }
}

function Write-GitHubOutput {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Value
    )

    if (-not $env:GITHUB_OUTPUT) {
        return
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::AppendAllText($env:GITHUB_OUTPUT, "$Name=$Value`n", $utf8NoBom)
}

Assert-SemanticVersion -Value $Version
Assert-FourPartVersion -Value $FileVersion -Name 'FileVersion'

$derivedInstallerVersion = Get-NumericInstallerVersion -Value $Version
if (-not $InstallerVersion) {
    $InstallerVersion = $derivedInstallerVersion
}
elseif ($InstallerVersion -ne $derivedInstallerVersion) {
    throw "InstallerVersion 必须由 Version 的数字部分推导：期望 $derivedInstallerVersion，实际 $InstallerVersion"
}
if ($InstallerVersion -notmatch '^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(?:\.(0|[1-9][0-9]*))?$') {
    throw "InstallerVersion 必须是两段或三段纯数字版本；实际值：$InstallerVersion"
}
Assert-VersionSegment -Value $InstallerVersion

if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path $installerRoot 'Output'
}
if (-not [IO.Path]::IsPathRooted($OutputDirectory)) {
    $OutputDirectory = Join-Path (Get-Location).Path $OutputDirectory
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path

if (-not $OutputBaseFilename) {
    $OutputBaseFilename = "AccuXSetup-$Version"
}
if ($OutputBaseFilename.EndsWith('.exe', [StringComparison]::OrdinalIgnoreCase)) {
    $OutputBaseFilename = $OutputBaseFilename.Substring(0, $OutputBaseFilename.Length - 4)
}
Assert-OutputBaseFilename -Value $OutputBaseFilename

if (-not (Test-Path -LiteralPath $Iscc -PathType Leaf)) {
    throw "未找到 Inno Setup 编译器：$Iscc"
}
if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
    throw "未找到安装脚本：$scriptPath"
}

$setupFileName = "$OutputBaseFilename.exe"
$setupPath = Join-Path $OutputDirectory $setupFileName
$checksumPath = "$setupPath.sha256"
$manifestPath = "$setupPath.manifest.json"
foreach ($stalePath in @($setupPath, $checksumPath, $manifestPath)) {
    if (Test-Path -LiteralPath $stalePath -PathType Leaf) {
        Remove-Item -LiteralPath $stalePath -Force
    }
}

$arguments = @(
    "/DAccuXDisplayVersion=$Version",
    "/DAccuXInstallerVersion=$InstallerVersion",
    "/DAccuXFileVersion=$FileVersion",
    "/DAccuXOutputBaseFilename=$OutputBaseFilename",
    "/O$OutputDirectory",
    $scriptPath
)
Write-Host "== 编译安装包：$Version（安装器 $InstallerVersion，文件 $FileVersion）==" -ForegroundColor Cyan
$compileOutput = & $Iscc @arguments 2>&1
$compileExitCode = $LASTEXITCODE
$compileOutput | ForEach-Object { Write-Host $_ }
if ($compileExitCode -ne 0) {
    throw "Inno Setup 编译失败，退出码 $compileExitCode。"
}

if (-not (Test-Path -LiteralPath $setupPath -PathType Leaf)) {
    throw "编译完成但未生成安装包：$setupPath"
}

$versionSummary = Assert-InstallerBinary `
    -Path $setupPath `
    -ExpectedFileVersion $FileVersion `
    -ExpectedDisplayVersion $Version

$hash = (Get-FileHash -LiteralPath $setupPath -Algorithm SHA256).Hash.ToLowerInvariant()
$fileName = [IO.Path]::GetFileName($setupPath)
$checksumText = "$hash *$fileName`r`n"
$ascii = [Text.Encoding]::ASCII
[IO.File]::WriteAllText($checksumPath, $checksumText, $ascii)

if (-not $CommitSha) {
    $CommitSha = if ($env:GITHUB_SHA) { $env:GITHUB_SHA } else { '' }
}
if (-not $RunNumber) {
    $RunNumber = if ($env:GITHUB_RUN_NUMBER) { $env:GITHUB_RUN_NUMBER } else { '' }
}
$manifest = [ordered]@{
    schemaVersion = 1
    version = $Version
    displayVersion = $Version
    installerVersion = $InstallerVersion
    fileVersion = $FileVersion
    fileName = $fileName
    sha256 = $hash
    sizeBytes = (Get-Item -LiteralPath $setupPath).Length
    commitSha = $CommitSha
    runNumber = $RunNumber
    runId = $RunNumber
    generatedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
}
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[IO.File]::WriteAllText($manifestPath, ($manifest | ConvertTo-Json -Depth 4), $utf8NoBom)

Write-GitHubOutput -Name 'setup_path' -Value $setupPath
Write-GitHubOutput -Name 'checksum_path' -Value $checksumPath
Write-GitHubOutput -Name 'manifest_path' -Value $manifestPath
Write-GitHubOutput -Name 'setup_filename' -Value $fileName
Write-GitHubOutput -Name 'checksum_filename' -Value ([IO.Path]::GetFileName($checksumPath))
Write-GitHubOutput -Name 'manifest_filename' -Value ([IO.Path]::GetFileName($manifestPath))
Write-GitHubOutput -Name 'display_version' -Value $Version
Write-GitHubOutput -Name 'installer_version' -Value $InstallerVersion
Write-GitHubOutput -Name 'file_version' -Value $FileVersion
Write-GitHubOutput -Name 'sha256' -Value $hash

Write-Host "安装包：$setupPath" -ForegroundColor Green
Write-Host "SHA-256：$hash" -ForegroundColor Green
Write-Host "版本资源：固定 FileVersion=$($versionSummary.FixedFileVersion)，文本=$($versionSummary.TextFileVersion)" -ForegroundColor Green
Write-Host "清单：$manifestPath" -ForegroundColor Green
