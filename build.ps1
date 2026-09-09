# AccuX 构建脚本：还原、构建、测试，并校验依赖边界。
#
# 用法：
#   powershell -ExecutionPolicy Bypass -File build.ps1
#   powershell -ExecutionPolicy Bypass -File build.ps1 -Configuration Release
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug',
    [switch]$SkipTests
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$solution = Join-Path $repoRoot 'AccuX.sln'

Write-Host "== 还原 ==" -ForegroundColor Cyan
& dotnet restore $solution
if ($LASTEXITCODE -ne 0) { throw '还原失败。' }

Write-Host "== 构建 ($Configuration) ==" -ForegroundColor Cyan
& dotnet build $solution -c $Configuration --no-restore --nologo
if ($LASTEXITCODE -ne 0) { throw '构建失败。' }

if (-not $SkipTests) {
    Write-Host '== 单元测试 ==' -ForegroundColor Cyan
    & dotnet test $solution -c $Configuration --no-build --nologo
    if ($LASTEXITCODE -ne 0) { throw '测试失败。' }
}

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
