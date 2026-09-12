using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using AccuX.AddIn.Updates;
using AccuX.Core.Logging;
using Xunit;

namespace AccuX.AddIn.Tests
{
    public sealed class UpdateServiceTests
    {
        [Fact]
        public async Task CheckLatest_ReturnsNewStableReleaseFromGitHubApi()
        {
            var client = CreateClient(_ => JsonResponse(ReleaseJson("1.4")));
            var service = new GitHubReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.True(result.IsSuccessful);
            Assert.True(result.HasUpdate);
            Assert.Equal("1.4", result.LatestRelease.Version.Text);
            Assert.Equal("https://github.com/smy116/AccuX/releases/tag/v1.4", result.LatestRelease.HtmlUrl);
            Assert.Equal("https://github.com/smy116/AccuX/releases/download/v1.4/AccuXSetup-1.4.exe", result.LatestRelease.InstallerUrl);
        }

        [Fact]
        public async Task CheckLatest_UsesGitHubLatestReleaseApi()
        {
            Uri requested = null;
            HttpRequestHeaders requestHeaders = null;
            var client = CreateClient(request =>
            {
                requested = request.RequestUri;
                requestHeaders = request.Headers;
                return JsonResponse(ReleaseJson("1.4"));
            });
            var service = new GitHubReleaseService(NullLogger.Instance, client);

            await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.Equal(GitHubReleaseService.LatestReleaseUrl, requested.AbsoluteUri);
            Assert.Equal("api.github.com", requested.Host);
            Assert.Contains(requestHeaders.UserAgent, value => value.Product != null && value.Product.Name == "AccuX-UpdateChecker");
            Assert.Contains(requestHeaders.GetValues("X-GitHub-Api-Version"), value => value == "2022-11-28");
        }

        [Fact]
        public async Task CheckLatest_FallsBackToGhProxyWhenGitHubIsUnavailable()
        {
            var requested = new System.Collections.Generic.List<string>();
            var client = CreateClient(request =>
            {
                requested.Add(request.RequestUri.AbsoluteUri);
                if (request.RequestUri.AbsoluteUri == GitHubReleaseService.LatestReleaseUrl)
                {
                    return new HttpResponseMessage(HttpStatusCode.BadGateway);
                }

                if (request.RequestUri.AbsoluteUri == GitHubReleaseService.ProxyLatestReleaseUrl)
                {
                    return JsonResponse(ReleaseJson("1.4"));
                }

                throw new InvalidOperationException("unexpected URL: " + request.RequestUri.AbsoluteUri);
            });
            var service = new GitHubReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.True(result.IsSuccessful);
            Assert.Equal(2, requested.Count);
            Assert.Equal(GitHubReleaseService.LatestReleaseUrl, requested[0]);
            Assert.Equal(GitHubReleaseService.ProxyLatestReleaseUrl, requested[1]);
            Assert.Equal(
                "https://gh-proxy.com/https://github.com/smy116/AccuX/releases/download/v1.4/AccuXSetup-1.4.exe",
                result.LatestRelease.InstallerUrl);
            Assert.Equal(
                "https://gh-proxy.com/https://github.com/smy116/AccuX/releases/tag/v1.4",
                result.LatestRelease.HtmlUrl);
        }

        [Fact]
        public void GetProxyUrl_PrefixesSupportedGitHubUrl()
        {
            Assert.Equal(
                "https://gh-proxy.com/https://github.com/smy116/AccuX/releases/download/v1.4/AccuXSetup-1.4.exe",
                GitHubReleaseService.GetProxyUrl("https://github.com/smy116/AccuX/releases/download/v1.4/AccuXSetup-1.4.exe"));
            Assert.Equal(
                "https://gh-proxy.com/https://api.github.com/repos/smy116/AccuX/releases/latest",
                GitHubReleaseService.GetProxyUrl("https://api.github.com/repos/smy116/AccuX/releases/latest"));
        }

        [Fact]
        public async Task CheckLatest_HttpErrorReturnsFailure()
        {
            var requests = 0;
            var client = CreateClient(_ =>
            {
                requests++;
                return new HttpResponseMessage(HttpStatusCode.NotFound);
            });
            var service = new GitHubReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.False(result.IsSuccessful);
            Assert.Contains("HTTP 404", result.ErrorMessage);
            Assert.Equal(1, requests);
        }

        [Theory]
        [InlineData("")]
        [InlineData("not-json")]
        public async Task CheckLatest_InvalidResponseReturnsFailure(string content)
        {
            var client = CreateClient(_ => JsonResponse(content));
            var service = new GitHubReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.False(result.IsSuccessful);
        }

        [Fact]
        public async Task CheckLatest_TimeoutReturnsFailure()
        {
            var client = new HttpClient(new BlockingHandler())
            {
                Timeout = TimeSpan.FromMilliseconds(50)
            };
            var service = new GitHubReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.False(result.IsSuccessful);
            Assert.Contains("超时", result.ErrorMessage);
        }

        [Theory]
        [InlineData("1.4", "1.4", false)]
        [InlineData("1.10", "1.4", false)]
        [InlineData("1.4", "1.10", true)]
        public async Task CheckLatest_ComparesVersionsNumerically(string currentVersion, string releaseVersion, bool expectedUpdate)
        {
            var client = CreateClient(_ => JsonResponse(ReleaseJson(releaseVersion)));
            var service = new GitHubReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync(currentVersion, CancellationToken.None);

            Assert.True(result.IsSuccessful);
            Assert.Equal(expectedUpdate, result.HasUpdate);
        }

        [Theory]
        [InlineData("v1.4-beta")]
        [InlineData("v1.4.1")]
        public async Task CheckLatest_RejectsUnsupportedOrMismatchedTag(string tag)
        {
            var client = CreateClient(_ => JsonResponse(ReleaseJson("1.4", tag)));
            var service = new GitHubReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.False(result.IsSuccessful);
        }

        [Fact]
        public async Task CheckLatest_RejectsInstallerUrlForDifferentTag()
        {
            var client = CreateClient(_ => JsonResponse(ReleaseJson(
                "1.4",
                "v1.5",
                installerUrl: "https://github.com/smy116/AccuX/releases/download/v1.4/AccuXSetup-1.5.exe")));
            var service = new GitHubReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.False(result.IsSuccessful);
        }

        [Fact]
        public async Task CheckLatest_RejectsDraftOrPrerelease()
        {
            var client = CreateClient(_ => JsonResponse(ReleaseJson("1.4", isDraft: true)));
            var service = new GitHubReleaseService(NullLogger.Instance, client);

            var draftResult = await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.False(draftResult.IsSuccessful);

            client = CreateClient(_ => JsonResponse(ReleaseJson("1.4", isPrerelease: true)));
            service = new GitHubReleaseService(NullLogger.Instance, client);
            var prereleaseResult = await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.False(prereleaseResult.IsSuccessful);
        }

        [Fact]
        public async Task CheckLatest_RejectsMissingInstaller()
        {
            var client = CreateClient(_ => JsonResponse(ReleaseJson("1.4", includeInstaller: false)));
            var service = new GitHubReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.False(result.IsSuccessful);
            Assert.Contains("安装包", result.ErrorMessage);
        }

        [Fact]
        public async Task CheckLatest_RejectsNonGitHubInstallerUrl()
        {
            var client = CreateClient(_ => JsonResponse(ReleaseJson(
                "1.4",
                installerUrl: "https://example.test/AccuXSetup-1.4.exe")));
            var service = new GitHubReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.False(result.IsSuccessful);
            Assert.Contains("安装包", result.ErrorMessage);
        }

        private static HttpClient CreateClient(Func<HttpRequestMessage, HttpResponseMessage> responder)
        {
            return new HttpClient(new StubHandler(responder));
        }

        private static string ReleaseJson(
            string version,
            string tag = null,
            bool isDraft = false,
            bool isPrerelease = false,
            bool includeInstaller = true,
            string installerUrl = null)
        {
            tag = tag ?? "v" + version;
            installerUrl = installerUrl ?? "https://github.com/smy116/AccuX/releases/download/" + tag + "/AccuXSetup-" + version + ".exe";
            var assets = includeInstaller
                ? ",\"assets\":[{\"name\":\"AccuXSetup-" + version + ".exe\",\"browser_download_url\":\"" + installerUrl + "\",\"size\":123}]"
                : ",\"assets\":[]";
            return "{\"tag_name\":\"" + tag + "\",\"name\":\"AccuX v" + version
                + "\",\"body\":\"修复问题\",\"html_url\":\"https://github.com/smy116/AccuX/releases/tag/" + tag
                + "\",\"published_at\":\"2026-01-01T00:00:00Z\",\"draft\":" + isDraft.ToString().ToLowerInvariant()
                + ",\"prerelease\":" + isPrerelease.ToString().ToLowerInvariant() + assets + "}";
        }

        private static HttpResponseMessage JsonResponse(string json)
        {
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(json, Encoding.UTF8, "application/json")
            };
        }

        private sealed class StubHandler : HttpMessageHandler
        {
            private readonly Func<HttpRequestMessage, HttpResponseMessage> _responder;

            public StubHandler(Func<HttpRequestMessage, HttpResponseMessage> responder)
            {
                _responder = responder;
            }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                return Task.FromResult(_responder(request));
            }
        }

        private sealed class BlockingHandler : HttpMessageHandler
        {
            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                await Task.Delay(Timeout.Infinite, cancellationToken);
                return null;
            }
        }
    }
}
