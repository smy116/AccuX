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
    /// 读取 jsDelivr 上的静态更新清单，并下载经过校验的安装包。
    ///
    /// 运行时只访问 cdn.jsdelivr.net。GitHub Release 仍由发布流程维护，
    /// 发布产物会同步到 jsDelivr 可读取的 update-feed 分支。
    /// </summary>
    internal sealed class JsDelivrReleaseService
    {
        private const string Repository = "smy116/AccuX";
        private const string FeedBranch = "update-feed";
        private const string LatestFeedPath = "latest.json";
        private const string CdnHost = "cdn.jsdelivr.net";
        private const string UserAgent = "AccuX-UpdateChecker";
        private const string LatestFeedUrl = "https://cdn.jsdelivr.net/gh/" + Repository + "@" + FeedBranch + "/" + LatestFeedPath;

        private static readonly Regex ChecksumLine = new Regex(
            @"^\s*([0-9a-fA-F]{64})\s+\*?(.+?)\s*$",
            RegexOptions.CultureInvariant);

        private readonly HttpClient _client;
        private readonly HttpClient _downloadClient;
        private readonly ILogger _logger;

        public JsDelivrReleaseService(ILogger logger, HttpClient client = null, HttpClient downloadClient = null)
        {
            _logger = logger ?? NullLogger.Instance;
            _client = client ?? CreateHttpClient(TimeSpan.FromSeconds(10));
            _downloadClient = downloadClient ?? (client == null ? CreateHttpClient(TimeSpan.FromMinutes(5)) : client);
        }

        /// <summary>
        /// 设置窗口的“查看更新说明”按钮也只打开 jsDelivr 地址。
        /// </summary>
        public static string ReleasePageUrl
        {
            get { return LatestFeedUrl; }
        }

        internal static string UpdateFeedUrl
        {
            get { return LatestFeedUrl; }
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
                    LatestFeedUrl,
                    cancellationToken).ConfigureAwait(false))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        return UpdateCheckResult.Failed(
                            currentVersion,
                            "jsDelivr 更新清单请求失败（HTTP " + (int)response.StatusCode + "）。");
                    }

                    var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    var dto = JsonConvert.DeserializeObject<JsDelivrUpdateFeedDto>(json);
                    if (dto == null || dto.Draft || dto.Prerelease)
                    {
                        return UpdateCheckResult.Failed(currentVersion, "jsDelivr 更新清单数据无效。");
                    }

                    if (!ReleaseVersion.TryParse(dto.Version, out var releaseVersion))
                    {
                        return UpdateCheckResult.Failed(currentVersion, "jsDelivr 更新清单版本号不是受支持的两段式稳定版本。");
                    }

                    var tagName = string.IsNullOrWhiteSpace(dto.Tag)
                        ? "v" + releaseVersion.Text
                        : dto.Tag.Trim();
                    if (!ReleaseVersion.TryParse(tagName, out var tagVersion)
                        || tagVersion.CompareTo(releaseVersion) != 0)
                    {
                        return UpdateCheckResult.Failed(currentVersion, "jsDelivr 更新清单中的 tag 与版本号不一致。");
                    }

                    var installerName = "AccuXSetup-" + releaseVersion.Text + ".exe";
                    var release = new UpdateReleaseInfo
                    {
                        TagName = tagName,
                        Name = string.IsNullOrWhiteSpace(dto.Name) ? "AccuX v" + releaseVersion.Text : dto.Name,
                        Body = dto.Notes ?? string.Empty,
                        HtmlUrl = IsJsDelivrUrl(dto.ReleaseNotesUrl) ? dto.ReleaseNotesUrl : LatestFeedUrl,
                        Draft = dto.Draft,
                        Prerelease = dto.Prerelease,
                        PublishedAtUtc = dto.PublishedAtUtc,
                        Version = releaseVersion,
                        Assets = new[]
                        {
                            new UpdateAsset { Name = installerName, DownloadUrl = dto.InstallerUrl },
                            new UpdateAsset { Name = installerName + ".sha256", DownloadUrl = dto.Sha256Url }
                        }
                    };

                    return UpdateCheckResult.Success(currentVersion, currentReleaseVersion, release);
                }
            }
            catch (TaskCanceledException ex) when (!cancellationToken.IsCancellationRequested)
            {
                _logger.Warn("jsDelivr 更新检测超时：" + ex.Message);
                return UpdateCheckResult.Failed(currentVersion, "jsDelivr 更新检测超时，请稍后重试。");
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                _logger.Warn("jsDelivr 更新检测失败：" + ex.Message);
                return UpdateCheckResult.Failed(currentVersion, "无法连接 jsDelivr 更新清单，请稍后重试。");
            }
        }

        public async Task<UpdateInstallResult> DownloadAndVerifyAsync(
            UpdateReleaseInfo release,
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
                var checksumText = await DownloadTextAsync(checksum.DownloadUrl, cancellationToken)
                    .ConfigureAwait(false);
                var expectedHash = ParseChecksum(checksumText, installerName);

                await DownloadFileAsync(installer.DownloadUrl, temporaryPath, progress, cancellationToken)
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

        internal bool HasRequiredAssets(UpdateReleaseInfo release, out string errorMessage)
        {
            return TryGetRequiredAssets(release, out _, out _, out errorMessage);
        }

        internal static string GetAssetUrl(string version, string fileName)
        {
            if (!ReleaseVersion.TryParse(version, out var releaseVersion)
                || string.IsNullOrWhiteSpace(fileName)
                || fileName.IndexOf('/') >= 0
                || fileName.IndexOf('\\') >= 0)
            {
                throw new ArgumentException("版本号或文件名无效。", nameof(version));
            }

            return "https://cdn.jsdelivr.net/gh/" + Repository + "@" + FeedBranch
                + "/releases/" + releaseVersion.Text + "/" + fileName;
        }

        private static HttpClient CreateHttpClient(TimeSpan timeout)
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var client = new HttpClient
            {
                Timeout = timeout
            };
            client.DefaultRequestHeaders.UserAgent.ParseAdd(UserAgent);
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
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

        private async Task<string> DownloadTextAsync(string url, CancellationToken cancellationToken)
        {
            using (var response = await _client.GetAsync(url, cancellationToken).ConfigureAwait(false))
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

        private static bool IsJsDelivrUrl(string value)
        {
            return Uri.TryCreate(value, UriKind.Absolute, out var uri)
                && string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)
                && string.Equals(uri.Host, CdnHost, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsJsDelivrUrl(string value, string expectedFileName)
        {
            if (!IsJsDelivrUrl(value) || string.IsNullOrWhiteSpace(expectedFileName))
            {
                return false;
            }

            var uri = new Uri(value, UriKind.Absolute);
            var path = uri.AbsolutePath.TrimEnd('/');
            var separator = path.LastIndexOf('/');
            var fileName = separator >= 0 ? path.Substring(separator + 1) : path;
            return string.Equals(fileName, expectedFileName, StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryGetRequiredAssets(
            UpdateReleaseInfo release,
            out UpdateAsset installer,
            out UpdateAsset checksum,
            out string errorMessage)
        {
            installer = null;
            checksum = null;
            errorMessage = null;
            if (release == null || release.Version == null)
            {
                errorMessage = "没有可用的更新信息。";
                return false;
            }

            var installerName = "AccuXSetup-" + release.Version.Text + ".exe";
            var checksumName = installerName + ".sha256";
            installer = release.FindAsset(installerName);
            checksum = release.FindAsset(checksumName);
            if (installer == null || checksum == null
                || !IsJsDelivrUrl(installer.DownloadUrl, installerName)
                || !IsJsDelivrUrl(checksum.DownloadUrl, checksumName))
            {
                errorMessage = "jsDelivr 更新清单缺少预期的安装包或 SHA-256 校验文件。";
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

        private sealed class JsDelivrUpdateFeedDto
        {
            [JsonProperty("version")]
            public string Version { get; set; }

            [JsonProperty("tag")]
            public string Tag { get; set; }

            [JsonProperty("name")]
            public string Name { get; set; }

            [JsonProperty("notes")]
            public string Notes { get; set; }

            [JsonProperty("releaseNotesUrl")]
            public string ReleaseNotesUrl { get; set; }

            [JsonProperty("installerUrl")]
            public string InstallerUrl { get; set; }

            [JsonProperty("sha256Url")]
            public string Sha256Url { get; set; }

            [JsonProperty("publishedAtUtc")]
            public DateTime? PublishedAtUtc { get; set; }

            [JsonProperty("draft")]
            public bool Draft { get; set; }

            [JsonProperty("prerelease")]
            public bool Prerelease { get; set; }
        }
    }
}
