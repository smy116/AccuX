<##
.SYNOPSIS
    Publishes the latest AccuX installer and update manifest to a jsDelivr-backed branch.
##>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Repository,

    [Parameter(Mandatory = $true)]
    [string]$Version,

    [Parameter(Mandatory = $true)]
    [string]$TagName,

    [Parameter(Mandatory = $true)]
    [string]$ArtifactDirectory,

    [string]$Branch = 'update-feed'
)

$ErrorActionPreference = 'Stop'

if ($Version -notmatch '^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)$') {
    throw "版本必须是两段数字且不能有前导零：$Version"
}

if ($Branch -notmatch '^[A-Za-z0-9._/-]+$') {
    throw "更新分支名称无效：$Branch"
}

$installerName = "AccuXSetup-$Version.exe"
$installerPath = Join-Path $ArtifactDirectory $installerName
$checksumPath = "$installerPath.sha256"
if (-not (Test-Path -LiteralPath $installerPath -PathType Leaf)) {
    throw "缺少安装包：$installerPath"
}
if (-not (Test-Path -LiteralPath $checksumPath -PathType Leaf)) {
    throw "缺少 SHA-256 文件：$checksumPath"
}

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("accux-jsdelivr-feed-" + [Guid]::NewGuid().ToString('N'))
$remoteUrl = "https://github.com/$Repository.git"
$cdnRoot = "https://cdn.jsdelivr.net/gh/$Repository@$Branch/releases/$Version"

function Invoke-Git {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Arguments)

    & git @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "git 命令失败：git $($Arguments -join ' ')"
    }
}

try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    Invoke-Git init $tempRoot
    Push-Location $tempRoot
    try {
        Invoke-Git remote add origin $remoteUrl

        # GitHub Actions 的 GH_TOKEN 由 gh 和 git 凭据配置共同使用。
        & gh auth setup-git
        if ($LASTEXITCODE -ne 0) {
            throw 'gh 未能配置 git 凭据。'
        }

        # PowerShell 会把 "$Branch:" 解析成特殊变量语法，必须先拆出完整 ref，
        # 否则实际传给 Git 的 refspec 会丢失分支名后的冒号和 refs 前缀。
        $branchRef = "refs/heads/$Branch"
        $trackingRef = "refs/remotes/origin/$Branch"
        $fetchRefspec = "${branchRef}:${trackingRef}"

        $remoteBranchOutput = & git ls-remote --exit-code --heads origin $branchRef 2>&1
        $remoteBranchExitCode = $LASTEXITCODE
        if ($remoteBranchExitCode -eq 0) {
            Invoke-Git fetch --depth=1 origin $fetchRefspec
            Invoke-Git checkout -B $Branch $trackingRef
        }
        elseif ($remoteBranchExitCode -eq 2) {
            Write-Host "未找到 $Branch 分支，将创建新的 jsDelivr 更新分支。"
            Invoke-Git checkout --orphan $Branch
        }
        else {
            $remoteBranchDetails = ($remoteBranchOutput -join [Environment]::NewLine)
            throw "无法检查远端 $Branch 分支：$remoteBranchDetails"
        }

        $releaseDirectory = Join-Path $tempRoot "releases\$Version"
        New-Item -ItemType Directory -Path $releaseDirectory -Force | Out-Null
        Copy-Item -LiteralPath $installerPath -Destination (Join-Path $releaseDirectory $installerName) -Force
        Copy-Item -LiteralPath $checksumPath -Destination (Join-Path $releaseDirectory "$installerName.sha256") -Force

        $releaseJson = & gh release view $TagName --repo $Repository --json name,body 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "无法读取 GitHub Release $TagName：$releaseJson"
        }
        $release = $releaseJson | ConvertFrom-Json
        $releaseName = if ([string]::IsNullOrWhiteSpace($release.name)) { "AccuX v$Version" } else { $release.name }
        $releaseBody = if ($null -eq $release.body) { '' } else { [string]$release.body }
        $notes = "# $releaseName`r`n`r`n$releaseBody"
        Set-Content -LiteralPath (Join-Path $releaseDirectory 'RELEASE-NOTES.md') -Value $notes -Encoding utf8

        $feed = [ordered]@{
            version = $Version
            tag = $TagName
            name = $releaseName
            notes = $releaseBody
            releaseNotesUrl = "$cdnRoot/RELEASE-NOTES.md"
            installerUrl = "$cdnRoot/$installerName"
            sha256Url = "$cdnRoot/$installerName.sha256"
            publishedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
            draft = $false
            prerelease = $false
        }
        $feedJson = $feed | ConvertTo-Json -Depth 4
        Set-Content -LiteralPath (Join-Path $tempRoot 'latest.json') -Value $feedJson -Encoding utf8

        Invoke-Git add --all
        & git diff --cached --quiet
        if ($LASTEXITCODE -eq 0) {
            Write-Host "jsDelivr 更新清单没有变化。"
            return
        }

        Invoke-Git config user.name 'github-actions[bot]'
        Invoke-Git config user.email '41898282+github-actions[bot]@users.noreply.github.com'
        Invoke-Git commit -m "chore: publish AccuX $Version jsDelivr feed"
        Invoke-Git push origin "HEAD:$branchRef"
        Write-Host "已发布 jsDelivr 更新清单：https://cdn.jsdelivr.net/gh/$Repository@$Branch/latest.json"
    }
    finally {
        Pop-Location
    }
}
finally {
    if (Test-Path -LiteralPath $tempRoot -PathType Container) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
