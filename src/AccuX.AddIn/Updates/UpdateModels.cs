using System;
using System.Collections.Generic;
using System.Linq;
using AccuX.Core.Updates;

namespace AccuX.AddIn.Updates
{
    internal sealed class UpdateReleaseInfo
    {
        public string TagName { get; set; }

        public string Name { get; set; }

        public string Body { get; set; }

        public string HtmlUrl { get; set; }

        public string InstallerUrl { get; set; }

        public bool Draft { get; set; }

        public bool Prerelease { get; set; }

        public DateTime? PublishedAtUtc { get; set; }

        public ReleaseVersion Version { get; set; }

        public IReadOnlyList<UpdateAsset> Assets { get; set; } = Array.Empty<UpdateAsset>();

        public UpdateAsset FindAsset(string name)
        {
            if (string.IsNullOrEmpty(name) || Assets == null)
            {
                return null;
            }

            return Assets.FirstOrDefault(asset =>
                asset != null && string.Equals(asset.Name, name, StringComparison.OrdinalIgnoreCase));
        }
    }

    internal sealed class UpdateAsset
    {
        public string Name { get; set; }

        public string DownloadUrl { get; set; }

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

        public UpdateReleaseInfo LatestRelease { get; private set; }

        public string ErrorMessage { get; private set; }

        public static UpdateCheckResult Success(
            string currentVersion,
            ReleaseVersion currentReleaseVersion,
            UpdateReleaseInfo latestRelease)
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

    internal sealed class UpdateOpenResult
    {
        private UpdateOpenResult()
        {
        }

        public bool Succeeded { get; private set; }

        public string Message { get; private set; }

        public static UpdateOpenResult Success()
        {
            return new UpdateOpenResult
            {
                Succeeded = true,
                Message = "已在浏览器中打开升级安装包下载地址。"
            };
        }

        public static UpdateOpenResult Failed(string message)
        {
            return new UpdateOpenResult
            {
                Succeeded = false,
                Message = message ?? "无法打开升级安装包下载地址。"
            };
        }
    }
}
