using System;
using System.Collections.Generic;
using System.Linq;
using AccuX.Core.Updates;

namespace AccuX.AddIn.Updates
{
    internal sealed class GitHubReleaseInfo
    {
        public string TagName { get; set; }

        public string Name { get; set; }

        public string Body { get; set; }

        public string HtmlUrl { get; set; }

        public bool Draft { get; set; }

        public bool Prerelease { get; set; }

        public DateTime? PublishedAtUtc { get; set; }

        public ReleaseVersion Version { get; set; }

        public IReadOnlyList<GitHubReleaseAsset> Assets { get; set; } = Array.Empty<GitHubReleaseAsset>();

        public GitHubReleaseAsset FindAsset(string name)
        {
            if (string.IsNullOrEmpty(name) || Assets == null)
            {
                return null;
            }

            return Assets.FirstOrDefault(asset =>
                asset != null && string.Equals(asset.Name, name, StringComparison.OrdinalIgnoreCase));
        }
    }

    internal sealed class GitHubReleaseAsset
    {
        public string Name { get; set; }

        public string BrowserDownloadUrl { get; set; }

        public long Size { get; set; }
    }

    internal sealed class UpdateCheckResult
    {
        private UpdateCheckResult()
        {
        }

        public bool IsSuccessful { get; private set; }

        public bool HasUpdate { get; private set; }

        public string CurrentVersion { get; private set; }

        public ReleaseVersion CurrentReleaseVersion { get; private set; }

        public GitHubReleaseInfo LatestRelease { get; private set; }

        public string ErrorMessage { get; private set; }

        public static UpdateCheckResult Success(
            string currentVersion,
            ReleaseVersion currentReleaseVersion,
            GitHubReleaseInfo latestRelease)
        {
            return new UpdateCheckResult
            {
                IsSuccessful = true,
                HasUpdate = latestRelease != null
                    && latestRelease.Version != null
                    && latestRelease.Version.CompareTo(currentReleaseVersion) > 0,
                CurrentVersion = currentVersion ?? string.Empty,
                CurrentReleaseVersion = currentReleaseVersion,
                LatestRelease = latestRelease
            };
        }

        public static UpdateCheckResult Failed(string currentVersion, string errorMessage)
        {
            return new UpdateCheckResult
            {
                IsSuccessful = false,
                CurrentVersion = currentVersion ?? string.Empty,
                ErrorMessage = errorMessage ?? "升级检测失败。"
            };
        }
    }

    internal sealed class UpdateInstallResult
    {
        private UpdateInstallResult()
        {
        }

        public bool Succeeded { get; private set; }

        public string Message { get; private set; }

        public string InstallerPath { get; private set; }

        public static UpdateInstallResult Success(string installerPath)
        {
            return new UpdateInstallResult
            {
                Succeeded = true,
                InstallerPath = installerPath,
                Message = "升级安装程序已启动。"
            };
        }

        public static UpdateInstallResult Failed(string message)
        {
            return new UpdateInstallResult
            {
                Succeeded = false,
                Message = message ?? "升级安装程序启动失败。"
            };
        }
    }
}
