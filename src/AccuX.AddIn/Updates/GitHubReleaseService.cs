using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using AccuX.Core.Logging;
using AccuX.Core.Updates;
using Newtonsoft.Json;

namespace AccuX.AddIn.Updates
{
    /// <summary>
    /// 读取 AccuX 的公开 GitHub Release，并下载经过校验的安装包。
    /// </summary>
    internal sealed class GitHubReleaseService
    {
        private const string Repository = "smy116/AccuX";
        private const string LatestReleaseUrl = "https://api.github.com/repos/" + Repository + "/releases/latest";
        private static readonly Regex ChecksumLine = new Regex(
            @"^\s*([0-9a-fA-F]{64})\s+\*?(.+?)\s*$",
            RegexOptions.CultureInvariant);

        private readonly HttpClient _client;
        private readonly HttpClient _downloadClient;
        private readonly ILogger _logger;

        public GitHubReleaseService(ILogger logger, HttpClient client = null, HttpClient downloadClient = null)
        {
            _logger = logger ?? NullLogger.Instance;
            _client = client ?? CreateHttpClient();
            _downloadClient = downloadClient ?? (client == null ? CreateDownloadHttpClient() : client);
        }

        public static string ReleasePageUrl
        {
            get { return "https://github.com/" + Repository + "/releases"; }
        }

        public async Task<UpdateCheckResult> CheckLatestAsync(
            string currentVersion,
            CancellationToken cancellationToken)
        {
            if (!TryParseCurrentVersion(currentVersion, out var currentReleaseVersion))
            {
                return UpdateCheckResult.Failed(currentVersion, "当前版本号无法识别，无法检测升级。");
            }

            try
            {
                using (var response = await _client.GetAsync(
                    LatestReleaseUrl,
                    cancellationToken).ConfigureAwait(false))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        return UpdateCheckResult.Failed(
                            currentVersion,
                            "GitHub Releases 请求失败（HTTP " + (int)response.StatusCode + "）。");
                    }

                    var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    var dto = JsonConvert.DeserializeObject<GitHubReleaseDto>(json);
                    if (dto == null || dto.Draft || dto.Prerelease)
                    {
                        return UpdateCheckResult.Failed(currentVersion, "GitHub Release 数据无效。");
                    }

                    if (!ReleaseVersion.TryParse(dto.TagName, out var releaseVersion))
                    {
                        return UpdateCheckResult.Failed(currentVersion, "GitHub Release 版本号不是受支持的两段式版本。");
                    }

                    var release = new GitHubReleaseInfo
                    {
                        TagName = dto.TagName,
                        Name = dto.Name,
                        Body = dto.Body,
                        HtmlUrl = dto.HtmlUrl,
                        Draft = dto.Draft,
                        Prerelease = dto.Prerelease,
                        PublishedAtUtc = dto.PublishedAt,
                        Version = releaseVersion,
                        Assets = ConvertAssets(dto.Assets)
                    };

                    return UpdateCheckResult.Success(currentVersion, currentReleaseVersion, release);
                }
            }
            catch (TaskCanceledException ex) when (!cancellationToken.IsCancellationRequested)
            {
                _logger.Warn("GitHub Release 检测超时：" + ex.Message);
                return UpdateCheckResult.Failed(currentVersion, "GitHub Releases 请求超时，请稍后重试。");
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                _logger.Warn("GitHub Release 检测失败：" + ex.Message);
                return UpdateCheckResult.Failed(currentVersion, "无法连接 GitHub Releases，请稍后重试。");
            }
        }

        public async Task<UpdateInstallResult> DownloadAndVerifyAsync(
            GitHubReleaseInfo release,
            IProgress<double> progress,
            CancellationToken cancellationToken)
        {
            if (!TryGetRequiredAssets(
                release,
                out var installer,
                out var checksum,
                out var assetError))
            {
                return UpdateInstallResult.Failed(assetError);
            }

            var installerName = "AccuXSetup-" + release.Version.Text + ".exe";

            var directory = Path.Combine(
                Path.GetTempPath(),
                "AccuX",
                "updates",
                release.Version.Text,
                Guid.NewGuid().ToString("N"));
            var temporaryPath = Path.Combine(directory, installerName + ".part");
            var finalPath = Path.Combine(directory, installerName);

            try
            {
                Directory.CreateDirectory(directory);
                var checksumText = await DownloadTextAsync(checksum.BrowserDownloadUrl, cancellationToken)
                    .ConfigureAwait(false);
                var expectedHash = ParseChecksum(checksumText, installerName);

                await DownloadFileAsync(installer.BrowserDownloadUrl, temporaryPath, progress, cancellationToken)
                    .ConfigureAwait(false);

                var actualHash = ComputeSha256(temporaryPath);
                if (!string.Equals(expectedHash, actualHash, StringComparison.OrdinalIgnoreCase))
                {
                    TryDelete(temporaryPath);
                    return UpdateInstallResult.Failed("安装包 SHA-256 校验失败，已取消启动。");
                }

                File.Move(temporaryPath, finalPath);
                return UpdateInstallResult.Success(finalPath);
            }
            catch (TaskCanceledException ex) when (!cancellationToken.IsCancellationRequested)
            {
                TryDelete(temporaryPath);
                _logger.Warn("升级安装包下载超时：" + ex.Message);
                return UpdateInstallResult.Failed("升级安装包下载超时，请稍后重试。");
            }
            catch (OperationCanceledException)
            {
                TryDelete(temporaryPath);
                throw;
            }
            catch (Exception ex)
            {
                TryDelete(temporaryPath);
                _logger.Warn("升级安装包下载失败：" + ex.Message);
                return UpdateInstallResult.Failed("升级安装包下载失败，请稍后重试。");
            }
        }

        internal bool HasRequiredAssets(GitHubReleaseInfo release, out string errorMessage)
        {
            return TryGetRequiredAssets(release, out _, out _, out errorMessage);
        }

        private static HttpClient CreateHttpClient()
        {
            return CreateHttpClient(TimeSpan.FromSeconds(10));
        }

        private static HttpClient CreateDownloadHttpClient()
        {
            return CreateHttpClient(TimeSpan.FromMinutes(5));
        }

        private static HttpClient CreateHttpClient(TimeSpan timeout)
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var client = new HttpClient
            {
                Timeout = timeout
            };
            client.DefaultRequestHeaders.UserAgent.ParseAdd("AccuX-UpdateChecker");
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
            return client;
        }

        private static bool TryParseCurrentVersion(string value, out ReleaseVersion version)
        {
            if (ReleaseVersion.TryParse(value, out version))
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var text = value.Trim();
            var plus = text.IndexOf('+');
            if (plus >= 0) text = text.Substring(0, plus);
            var dash = text.IndexOf('-');
            if (dash >= 0) text = text.Substring(0, dash);
            var parts = text.TrimStart('v', 'V').Split('.');
            if (parts.Length < 2)
            {
                return false;
            }

            return ReleaseVersion.TryParse(parts[0] + "." + parts[1], out version);
        }

        private static IReadOnlyList<GitHubReleaseAsset> ConvertAssets(GitHubReleaseAssetDto[] assets)
        {
            if (assets == null || assets.Length == 0)
            {
                return Array.Empty<GitHubReleaseAsset>();
            }

            var result = new GitHubReleaseAsset[assets.Length];
            for (var i = 0; i < assets.Length; i++)
            {
                var asset = assets[i];
                result[i] = asset == null
                    ? null
                    : new GitHubReleaseAsset
                    {
                        Name = asset.Name,
                        BrowserDownloadUrl = asset.BrowserDownloadUrl,
                        Size = asset.Size
                    };
            }

            return result;
        }

        private async Task<string> DownloadTextAsync(string url, CancellationToken cancellationToken)
        {
            using (var response = await _client.GetAsync(
                url,
                cancellationToken).ConfigureAwait(false))
            {
                if (!response.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException("HTTP " + (int)response.StatusCode);
                }

                return await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            }
        }

        private async Task DownloadFileAsync(
            string url,
            string path,
            IProgress<double> progress,
            CancellationToken cancellationToken)
        {
            using (var requestCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
            {
                if (_downloadClient.Timeout != System.Threading.Timeout.InfiniteTimeSpan)
                {
                    requestCancellation.CancelAfter(_downloadClient.Timeout);
                }

                await DownloadFileCoreAsync(url, path, progress, requestCancellation.Token).ConfigureAwait(false);
            }
        }

        private async Task DownloadFileCoreAsync(
            string url,
            string path,
            IProgress<double> progress,
            CancellationToken cancellationToken)
        {
            using (var response = await _downloadClient.GetAsync(
                url,
                HttpCompletionOption.ResponseHeadersRead,
                cancellationToken).ConfigureAwait(false))
            {
                if (!response.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException("HTTP " + (int)response.StatusCode);
                }

                var total = response.Content.Headers.ContentLength;
                var buffer = new byte[64 * 1024];
                long completed = 0;
                using (var input = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None, buffer.Length, true))
                {
                    int read;
                    while ((read = await input.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false)) > 0)
                    {
                        await output.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
                        completed += read;
                        if (total.HasValue && total.Value > 0)
                        {
                            progress?.Report(completed * 100d / total.Value);
                        }
                    }
                }
            }
        }

        private static string ParseChecksum(string text, string installerName)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                throw new InvalidOperationException("SHA-256 文件为空。");
            }

            foreach (var line in text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var match = ChecksumLine.Match(line);
                if (!match.Success)
                {
                    continue;
                }

                var fileName = match.Groups[2].Value.Trim().Trim('"');
                if (!string.Equals(fileName, installerName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                return match.Groups[1].Value.ToLowerInvariant();
            }

            throw new InvalidOperationException("SHA-256 文件与安装包名称不匹配。");
        }

        private static string ComputeSha256(string path)
        {
            using (var sha256 = SHA256.Create())
            using (var stream = File.OpenRead(path))
            {
                var hash = sha256.ComputeHash(stream);
                return BitConverter.ToString(hash).Replace("-", string.Empty).ToLowerInvariant();
            }
        }

        private static bool IsHttpsUrl(string value)
        {
            return Uri.TryCreate(value, UriKind.Absolute, out var uri)
                && string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryGetRequiredAssets(
            GitHubReleaseInfo release,
            out GitHubReleaseAsset installer,
            out GitHubReleaseAsset checksum,
            out string errorMessage)
        {
            installer = null;
            checksum = null;
            errorMessage = null;
            if (release == null || release.Version == null)
            {
                errorMessage = "没有可用的 Release 信息。";
                return false;
            }

            var installerName = "AccuXSetup-" + release.Version.Text + ".exe";
            var checksumName = installerName + ".sha256";
            installer = release.FindAsset(installerName);
            checksum = release.FindAsset(checksumName);
            if (installer == null || checksum == null
                || !IsHttpsUrl(installer.BrowserDownloadUrl)
                || !IsHttpsUrl(checksum.BrowserDownloadUrl))
            {
                errorMessage = "Release 缺少预期的安装包或 SHA-256 校验文件，请打开发布页手动查看。";
                return false;
            }

            return true;
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
                // 临时文件清理失败不应覆盖原始错误。
            }
        }

        private sealed class GitHubReleaseDto
        {
            [JsonProperty("tag_name")]
            public string TagName { get; set; }

            [JsonProperty("name")]
            public string Name { get; set; }

            [JsonProperty("body")]
            public string Body { get; set; }

            [JsonProperty("html_url")]
            public string HtmlUrl { get; set; }

            [JsonProperty("draft")]
            public bool Draft { get; set; }

            [JsonProperty("prerelease")]
            public bool Prerelease { get; set; }

            [JsonProperty("published_at")]
            public DateTime? PublishedAt { get; set; }

            [JsonProperty("assets")]
            public GitHubReleaseAssetDto[] Assets { get; set; }
        }

        private sealed class GitHubReleaseAssetDto
        {
            [JsonProperty("name")]
            public string Name { get; set; }

            [JsonProperty("browser_download_url")]
            public string BrowserDownloadUrl { get; set; }

            [JsonProperty("size")]
            public long Size { get; set; }
        }
    }
}
