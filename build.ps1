# AccuX 构建脚本：还原、构建、测试，并校验依赖边界。
#
# 用法：
#   powershell -ExecutionPolicy Bypass -File build.ps1
#   powershell -ExecutionPolicy Bypass -File build.ps1 -Configuration Release -Version 1.1
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug',
    [string]$Version = '1.1',
    [string]$OfficePiaPath = '',
    [switch]$SkipTests
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$solution = Join-Path $repoRoot 'AccuX.sln'

function Assert-TwoPartVersion {
    param([Parameter(Mandatory)][string]$Value)

    if ($Value -notmatch '^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)$') {
        throw "版本必须是两段数字且不能有前导零，例如 1.1；实际值：$Value"
    }

    $parts = $Value.Split('.')
    foreach ($part in $parts) {
        $number = [int64]$part
        if ($number -gt 65534) {
            throw "版本段必须在 0 到 65534 之间：$Value"
        }
    }
}

function Get-PiaAssemblyVersion {
    param([Parameter(Mandatory)][string]$Path)

    try {
        return [Reflection.AssemblyName]::GetAssemblyName($Path).Version.ToString()
    }
    catch {
        throw "无法读取 Office PIA 版本：$Path。$($_.Exception.Message)"
    }
}

function Assert-OfficePiaFiles {
    param([Parameter(Mandatory)][string[]]$Paths)

    foreach ($path in $Paths) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "缺少 Office PIA：$path"
        }

        $version = Get-PiaAssemblyVersion -Path $path
        if ($version -ne '15.0.0.0') {
            throw "Office PIA 必须是 15.0.0.0：$path 实际为 $version"
        }
    }
}

function Resolve-OfficePia {
    param([string]$ExplicitPath)

    $fileNames = @('Microsoft.Office.Interop.Excel.dll', 'Office.dll', 'Microsoft.Vbe.Interop.dll')
    if ($ExplicitPath) {
        $resolved = (Resolve-Path -LiteralPath $ExplicitPath -ErrorAction Stop).Path
        $paths = $fileNames | ForEach-Object { Join-Path $resolved $_ }
        Assert-OfficePiaFiles -Paths $paths
        return $resolved
    }

    $candidates = @()
    if (${env:ProgramFiles(x86)}) {
        $candidates += Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15'
    }
    if ($env:ProgramFiles) {
        $candidates += Join-Path $env:ProgramFiles 'Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15'
    }

    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        if (-not (Test-Path -LiteralPath $candidate -PathType Container)) {
            continue
        }

        $paths = $fileNames | ForEach-Object { Join-Path $candidate $_ }
        if (($paths | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) }).Count -eq 0) {
            Assert-OfficePiaFiles -Paths $paths
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    # Older developer machines may only expose the same PIAs through the GAC.
    $gacPaths = @(
        (Join-Path $env:WINDIR 'assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Excel.dll'),
        (Join-Path $env:WINDIR 'assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll'),
        (Join-Path $env:WINDIR 'assembly\GAC_MSIL\Microsoft.Vbe.Interop\15.0.0.0__71e9bce111e9429c\Microsoft.Vbe.Interop.dll')
    )
    Assert-OfficePiaFiles -Paths $gacPaths
    return $null
}

Assert-TwoPartVersion -Value $Version
$versionParts = $Version.Split('.')
$assemblyVersion = "$($versionParts[0]).$($versionParts[1]).0.0"
$resolvedOfficePiaPath = Resolve-OfficePia -ExplicitPath $OfficePiaPath
$msbuildProperties = @(
    "-p:Version=$Version",
    "-p:AccuXVersion=$Version",
    "-p:AssemblyVersion=$assemblyVersion",
    "-p:FileVersion=$assemblyVersion"
)
if ($resolvedOfficePiaPath) {
    $msbuildProperties += "-p:OfficePiaPath=$resolvedOfficePiaPath"
}

Write-Host "== 版本：$Version（程序集 $assemblyVersion）==" -ForegroundColor Cyan
if ($resolvedOfficePiaPath) {
    Write-Host "== Office PIA：$resolvedOfficePiaPath ==" -ForegroundColor Cyan
}

Write-Host "== 还原 ==" -ForegroundColor Cyan
& dotnet restore $solution @msbuildProperties
if ($LASTEXITCODE -ne 0) { throw '还原失败。' }

Write-Host "== 构建 ($Configuration) ==" -ForegroundColor Cyan
& dotnet build $solution -c $Configuration --no-restore --nologo @msbuildProperties
if ($LASTEXITCODE -ne 0) { throw '构建失败。' }

if (-not $SkipTests) {
    Write-Host '== 单元测试 ==' -ForegroundColor Cyan
    & dotnet test $solution -c $Configuration --no-build --nologo @msbuildProperties
    if ($LASTEXITCODE -ne 0) { throw '测试失败。' }
}

$outputDirectory = Join-Path $repoRoot "src\AccuX.AddIn\bin\$Configuration\net48"
$outputPiaFiles = @(
    (Join-Path $outputDirectory 'Microsoft.Office.Interop.Excel.dll'),
    (Join-Path $outputDirectory 'office.dll'),
    (Join-Path $outputDirectory 'Microsoft.Vbe.Interop.dll')
)
Assert-OfficePiaFiles -Paths $outputPiaFiles
Write-Host 'Office PIA 校验通过：三个 15.0.0.0 程序集已复制到输出目录。' -ForegroundColor Green

Write-Host '== 依赖边界校验（编码规则 1：Core 不引用 Excel/WPS Interop）==' -ForegroundColor Cyan
$coreDll = Join-Path $repoRoot "src\AccuX.Core\bin\$Configuration\net48\AccuX.Core.dll"
$assembly = [System.Reflection.Assembly]::LoadFrom($coreDll)
$forbidden = $assembly.GetReferencedAssemblies() |
    Where-Object { $_.Name -match 'Interop|Excel|Office|Wps|Kingsoft' }

if ($forbidden) {
    $names = ($forbidden | ForEach-Object { $_.Name }) -join ', '
    throw "依赖边界违规：AccuX.Core 引用了 $names"
}

Write-Host '依赖边界校验通过：AccuX.Core 未引用任何 Excel/WPS Interop。' -ForegroundColor Green

# 校验 Core 也不引用 Host 工程。
$hostRef = $assembly.GetReferencedAssemblies() | Where-Object { $_.Name -eq 'AccuX.Host' }
if ($hostRef) {
    throw '依赖边界违规：AccuX.Core 引用了 AccuX.Host。'
}

Write-Host '依赖边界校验通过：AccuX.Core 未引用 AccuX.Host。' -ForegroundColor Green
Write-Host "`n全部完成。输出目录：src\*\bin\$Configuration\net48" -ForegroundColor Green
