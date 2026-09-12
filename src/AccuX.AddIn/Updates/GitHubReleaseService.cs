using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using AccuX.Core.Logging;
using AccuX.Core.Updates;
using Newtonsoft.Json;

namespace AccuX.AddIn.Updates
{
    /// <summary>
    /// Reads the latest stable GitHub Release and exposes its installer URL.
    /// The installer is intentionally opened in the user's browser; AccuX does
    /// not download, verify, or launch the installer itself.
    /// </summary>
    internal sealed class GitHubReleaseService
    {
        private const string Repository = "smy116/AccuX";
        private const string GitHubHost = "github.com";
        private const string ApiHost = "api.github.com";
        private const string UserAgent = "AccuX-UpdateChecker";
        private const string ProxyBaseUrl = "https://gh-proxy.com/";
        private const string LatestReleaseApiUrl = "https://api.github.com/repos/" + Repository + "/releases/latest";

        private readonly HttpClient _client;
        private readonly ILogger _logger;

        public GitHubReleaseService(ILogger logger, HttpClient client = null)
        {
            _logger = logger ?? NullLogger.Instance;
            _client = client ?? CreateHttpClient(TimeSpan.FromSeconds(10));
            ConfigureHttpClient(_client);
        }

        internal static string LatestReleaseUrl
        {
            get { return LatestReleaseApiUrl; }
        }

        internal static string ProxyLatestReleaseUrl
        {
            get { return ToProxyUrl(LatestReleaseApiUrl); }
        }

        internal static string GetProxyUrl(string githubUrl)
        {
            return ToProxyUrl(githubUrl);
        }

        public async Task<UpdateCheckResult> CheckLatestAsync(
            string currentVersion,
            CancellationToken cancellationToken)
        {
            if (!TryParseCurrentVersion(currentVersion, out var currentReleaseVersion))
            {
                return UpdateCheckResult.Failed(currentVersion, "当前版本号无法识别，无法检测升级。");
            }

            var direct = await TryCheckLatestAsync(
                currentVersion,
                currentReleaseVersion,
                LatestReleaseApiUrl,
                false,
                cancellationToken).ConfigureAwait(false);
            if (!direct.TryProxy)
            {
                return direct.Result;
            }

            _logger.Warn("GitHub Release 直连失败，尝试 gh-proxy.com：" + direct.Result.ErrorMessage);
            var proxy = await TryCheckLatestAsync(
                currentVersion,
                currentReleaseVersion,
                ProxyLatestReleaseUrl,
                true,
                cancellationToken).ConfigureAwait(false);
            return proxy.Result;
        }

        internal bool HasInstaller(UpdateReleaseInfo release, out string errorMessage)
        {
            errorMessage = null;
            if (release == null || release.Version == null || string.IsNullOrWhiteSpace(release.InstallerUrl))
            {
                errorMessage = "GitHub Release 缺少有效的安装包附件。";
                return false;
            }

            var installerName = "AccuXSetup-" + release.Version.Text + ".exe";
            if (!IsSupportedAssetUrl(release.InstallerUrl, release.TagName, release.Version, installerName))
            {
                errorMessage = "GitHub Release 安装包地址无效。";
                return false;
            }

            return true;
        }

        private async Task<FetchResult> TryCheckLatestAsync(
            string currentVersion,
            ReleaseVersion currentReleaseVersion,
            string endpoint,
            bool useProxy,
            CancellationToken cancellationToken)
        {
            try
            {
                using (var response = await _client.GetAsync(
                    endpoint,
                    cancellationToken).ConfigureAwait(false))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        var failed = UpdateCheckResult.Failed(
                            currentVersion,
                            "GitHub Release 请求失败（HTTP " + (int)response.StatusCode + "）。");
                        return ShouldTryProxy(response.StatusCode)
                            ? FetchResult.Retryable(failed)
                            : FetchResult.Invalid(failed);
                    }

                    var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    var dto = JsonConvert.DeserializeObject<GitHubReleaseDto>(json);
                    if (dto == null || dto.Draft || dto.Prerelease)
                    {
                        return FetchResult.Invalid(UpdateCheckResult.Failed(currentVersion, "GitHub Release 数据无效。"));
                    }

                    if (!ReleaseVersion.TryParse(dto.TagName, out var releaseVersion))
                    {
                        return FetchResult.Invalid(UpdateCheckResult.Failed(
                            currentVersion,
                            "GitHub Release tag 不是受支持的两段式稳定版本。"));
                    }

                    var installerName = "AccuXSetup-" + releaseVersion.Text + ".exe";
                    var installer = FindInstaller(dto.Assets, installerName, dto.TagName, releaseVersion);
                    if (installer == null)
                    {
                        return FetchResult.Invalid(UpdateCheckResult.Failed(
                            currentVersion,
                            "GitHub Release 缺少有效的安装包附件。"));
                    }

                    if (!IsGitHubReleasePageUrl(dto.HtmlUrl, dto.TagName))
                    {
                        return FetchResult.Invalid(UpdateCheckResult.Failed(
                            currentVersion,
                            "GitHub Release 缺少有效的更新说明地址。"));
                    }

                    var release = new UpdateReleaseInfo
                    {
                        TagName = dto.TagName.Trim(),
                        Name = string.IsNullOrWhiteSpace(dto.Name) ? "AccuX v" + releaseVersion.Text : dto.Name,
                        Body = dto.Body ?? string.Empty,
                        HtmlUrl = useProxy ? ToProxyUrl(dto.HtmlUrl) : dto.HtmlUrl,
                        InstallerUrl = useProxy ? ToProxyUrl(installer.BrowserDownloadUrl) : installer.BrowserDownloadUrl,
                        Draft = dto.Draft,
                        Prerelease = dto.Prerelease,
                        PublishedAtUtc = dto.PublishedAtUtc,
                        Version = releaseVersion,
                        Assets = ConvertAssets(dto.Assets, useProxy)
                    };

                    return FetchResult.Success(UpdateCheckResult.Success(currentVersion, currentReleaseVersion, release));
                }
            }
            catch (TaskCanceledException ex) when (!cancellationToken.IsCancellationRequested)
            {
                _logger.Warn("GitHub Release 检测超时：" + ex.Message);
                return FetchResult.Retryable(UpdateCheckResult.Failed(
                    currentVersion,
                    "GitHub Release 检测超时，请稍后重试。"));
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (JsonException ex)
            {
                _logger.Warn("GitHub Release 响应 JSON 无效：" + ex.Message);
                return FetchResult.Invalid(UpdateCheckResult.Failed(
                    currentVersion,
                    "GitHub Release 响应数据无效。"));
            }
            catch (Exception ex)
            {
                _logger.Warn("GitHub Release 检测失败：" + ex.Message);
                return FetchResult.Retryable(UpdateCheckResult.Failed(
                    currentVersion,
                    "无法连接 GitHub Release，请稍后重试。"));
            }
        }

        private static HttpClient CreateHttpClient(TimeSpan timeout)
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var client = new HttpClient
            {
                Timeout = timeout
            };
            return client;
        }

        private static void ConfigureHttpClient(HttpClient client)
        {
            if (client == null)
            {
                return;
            }

            client.DefaultRequestHeaders.UserAgent.ParseAdd(UserAgent);
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
            client.DefaultRequestHeaders.TryAddWithoutValidation("X-GitHub-Api-Version", "2022-11-28");
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

        private static GitHubAssetDto FindInstaller(
            IReadOnlyList<GitHubAssetDto> assets,
            string installerName,
            string tagName,
            ReleaseVersion version)
        {
            if (assets == null)
            {
                return null;
            }

            foreach (var asset in assets)
            {
                if (asset == null
                    || !string.Equals(asset.Name, installerName, StringComparison.OrdinalIgnoreCase)
                    || !IsGitHubAssetUrl(asset.BrowserDownloadUrl, tagName, version, installerName))
                {
                    continue;
                }

                return asset;
            }

            return null;
        }

        private static IReadOnlyList<UpdateAsset> ConvertAssets(
            IReadOnlyList<GitHubAssetDto> assets,
            bool useProxy)
        {
            var result = new List<UpdateAsset>();
            if (assets == null)
            {
                return result;
            }

            foreach (var asset in assets)
            {
                if (asset == null)
                {
                    continue;
                }

                result.Add(new UpdateAsset
                {
                    Name = asset.Name,
                    DownloadUrl = useProxy ? ToProxyUrl(asset.BrowserDownloadUrl) : asset.BrowserDownloadUrl,
                    Size = asset.Size
                });
            }

            return result;
        }

        private static bool IsGitHubReleasePageUrl(string value, string tagName)
        {
            if (!Uri.TryCreate(value, UriKind.Absolute, out var uri)
                || !string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(uri.Host, GitHubHost, StringComparison.OrdinalIgnoreCase)
                || string.IsNullOrWhiteSpace(tagName))
            {
                return false;
            }

            var expectedPath = "/" + Repository + "/releases/tag/" + tagName.Trim();
            return string.Equals(uri.AbsolutePath, expectedPath, StringComparison.OrdinalIgnoreCase)
                && string.IsNullOrEmpty(uri.Query)
                && string.IsNullOrEmpty(uri.Fragment);
        }

        private static bool IsSupportedAssetUrl(
            string value,
            string tagName,
            ReleaseVersion version,
            string expectedFileName)
        {
            var direct = RemoveProxyPrefix(value);
            return IsGitHubAssetUrl(direct, tagName, version, expectedFileName);
        }

        private static bool IsGitHubAssetUrl(
            string value,
            string tagName,
            ReleaseVersion version,
            string expectedFileName)
        {
            if (!Uri.TryCreate(value, UriKind.Absolute, out var uri)
                || !string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(uri.Host, GitHubHost, StringComparison.OrdinalIgnoreCase)
                || version == null
                || !ReleaseVersion.TryParse(tagName, out var tagVersion)
                || tagVersion.CompareTo(version) != 0
                || !IsSafeFileName(expectedFileName))
            {
                return false;
            }

            var expectedPath = "/" + Repository + "/releases/download/" + tagName.Trim() + "/" + expectedFileName;
            return string.Equals(uri.AbsolutePath, expectedPath, StringComparison.OrdinalIgnoreCase)
                && string.IsNullOrEmpty(uri.Query)
                && string.IsNullOrEmpty(uri.Fragment);
        }

        private static string ToProxyUrl(string githubUrl)
        {
            if (string.IsNullOrWhiteSpace(githubUrl)
                || githubUrl.StartsWith(ProxyBaseUrl, StringComparison.OrdinalIgnoreCase))
            {
                return githubUrl;
            }

            if (!Uri.TryCreate(githubUrl, UriKind.Absolute, out var uri)
                || !string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)
                || (!string.Equals(uri.Host, GitHubHost, StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(uri.Host, ApiHost, StringComparison.OrdinalIgnoreCase)))
            {
                return githubUrl;
            }

            return ProxyBaseUrl + uri.AbsoluteUri;
        }

        private static string RemoveProxyPrefix(string value)
        {
            if (!string.IsNullOrWhiteSpace(value)
                && value.StartsWith(ProxyBaseUrl, StringComparison.OrdinalIgnoreCase))
            {
                return value.Substring(ProxyBaseUrl.Length);
            }

            return value;
        }

        private static bool IsSafeFileName(string value)
        {
            return !string.IsNullOrWhiteSpace(value)
                && value.IndexOf('/') < 0
                && value.IndexOf('\\') < 0;
        }

        private static bool ShouldTryProxy(HttpStatusCode statusCode)
        {
            return statusCode == HttpStatusCode.RequestTimeout
                || statusCode == HttpStatusCode.Forbidden
                || (int)statusCode == 429
                || (int)statusCode >= 500;
        }

        private sealed class FetchResult
        {
            private FetchResult(UpdateCheckResult result, bool tryProxy)
            {
                Result = result;
                TryProxy = tryProxy;
            }

            public UpdateCheckResult Result { get; }

            public bool TryProxy { get; }

            public static FetchResult Success(UpdateCheckResult result)
            {
                return new FetchResult(result, false);
            }

            public static FetchResult Invalid(UpdateCheckResult result)
            {
                return new FetchResult(result, false);
            }

            public static FetchResult Retryable(UpdateCheckResult result)
            {
                return new FetchResult(result, true);
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

            [JsonProperty("published_at")]
            public DateTime? PublishedAtUtc { get; set; }

            [JsonProperty("draft")]
            public bool Draft { get; set; }

            [JsonProperty("prerelease")]
            public bool Prerelease { get; set; }

            [JsonProperty("assets")]
            public List<GitHubAssetDto> Assets { get; set; }
        }

        private sealed class GitHubAssetDto
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
