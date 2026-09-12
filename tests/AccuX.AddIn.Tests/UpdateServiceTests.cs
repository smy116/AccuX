using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
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
        public async Task CheckLatest_ReturnsNewStableReleaseFromJsDelivrFeed()
        {
            var client = CreateClient(request => JsonResponse(FeedJson("1.4")));
            var service = new JsDelivrReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.True(result.IsSuccessful);
            Assert.True(result.HasUpdate);
            Assert.Equal("1.4", result.LatestRelease.Version.Text);
            Assert.Contains("cdn.jsdelivr.net", result.LatestRelease.HtmlUrl);
        }

        [Fact]
        public async Task CheckLatest_UsesOnlyJsDelivrFeedUrl()
        {
            Uri requested = null;
            var client = CreateClient(request =>
            {
                requested = request.RequestUri;
                return JsonResponse(FeedJson("1.4"));
            });
            var service = new JsDelivrReleaseService(NullLogger.Instance, client);

            await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.Equal(JsDelivrReleaseService.UpdateFeedUrl, requested.AbsoluteUri);
            Assert.Equal("cdn.jsdelivr.net", requested.Host);
        }

        [Fact]
        public async Task CheckLatest_HttpErrorReturnsFailure()
        {
            var client = CreateClient(_ => new HttpResponseMessage(HttpStatusCode.NotFound));
            var service = new JsDelivrReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.False(result.IsSuccessful);
            Assert.Contains("HTTP 404", result.ErrorMessage);
        }

        [Theory]
        [InlineData("1.4", "1.4", false)]
        [InlineData("1.10", "1.4", false)]
        [InlineData("1.4", "1.10", true)]
        public async Task CheckLatest_ComparesVersionsNumerically(string currentVersion, string feedVersion, bool expectedUpdate)
        {
            var client = CreateClient(_ => JsonResponse(FeedJson(feedVersion)));
            var service = new JsDelivrReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync(currentVersion, CancellationToken.None);

            Assert.True(result.IsSuccessful);
            Assert.Equal(expectedUpdate, result.HasUpdate);
        }

        [Theory]
        [InlineData("1.4-beta")]
        [InlineData("1.4.1")]
        public async Task CheckLatest_RejectsPrereleaseOrUnsupportedVersion(string version)
        {
            var client = CreateClient(_ => JsonResponse(FeedJson(version)));
            var service = new JsDelivrReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.False(result.IsSuccessful);
        }

        [Fact]
        public async Task CheckLatest_RejectsMismatchedTag()
        {
            var client = CreateClient(_ => JsonResponse(FeedJson("1.4", tag: "v1.5")));
            var service = new JsDelivrReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.False(result.IsSuccessful);
            Assert.Contains("tag", result.ErrorMessage);
        }

        [Fact]
        public async Task DownloadAndVerify_RequiresMatchingJsDelivrChecksum()
        {
            var installerName = "AccuXSetup-1.4.exe";
            var installerUrl = JsDelivrReleaseService.GetAssetUrl("1.4", installerName);
            var checksumUrl = JsDelivrReleaseService.GetAssetUrl("1.4", installerName + ".sha256");
            var installerBytes = Encoding.UTF8.GetBytes("test installer");
            var hash = ComputeSha256(installerBytes);
            var responses = new Dictionary<string, HttpResponseMessage>(StringComparer.Ordinal)
            {
                [JsDelivrReleaseService.UpdateFeedUrl] = JsonResponse(FeedJson("1.4", installerUrl, checksumUrl)),
                [checksumUrl] = TextResponse(hash + " *" + installerName),
                [installerUrl] = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ByteArrayContent(installerBytes)
                }
            };
            var client = CreateClient(request => responses[request.RequestUri.AbsoluteUri]);
            var service = new JsDelivrReleaseService(NullLogger.Instance, client);
            var check = await service.CheckLatestAsync("1.3", CancellationToken.None);

            var result = await service.DownloadAndVerifyAsync(check.LatestRelease, null, CancellationToken.None);

            Assert.True(result.Succeeded);
            Assert.True(System.IO.File.Exists(result.InstallerPath));
            System.IO.File.Delete(result.InstallerPath);
            System.IO.Directory.Delete(System.IO.Path.GetDirectoryName(result.InstallerPath));
        }

        [Fact]
        public async Task DownloadAndVerify_RejectsWrongChecksum()
        {
            var installerName = "AccuXSetup-1.4.exe";
            var installerUrl = JsDelivrReleaseService.GetAssetUrl("1.4", installerName);
            var checksumUrl = JsDelivrReleaseService.GetAssetUrl("1.4", installerName + ".sha256");
            var responses = new Dictionary<string, HttpResponseMessage>(StringComparer.Ordinal)
            {
                [JsDelivrReleaseService.UpdateFeedUrl] = JsonResponse(FeedJson("1.4", installerUrl, checksumUrl)),
                [checksumUrl] = TextResponse(new string('0', 64) + " *" + installerName),
                [installerUrl] = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ByteArrayContent(Encoding.UTF8.GetBytes("test installer"))
                }
            };
            var client = CreateClient(request => responses[request.RequestUri.AbsoluteUri]);
            var service = new JsDelivrReleaseService(NullLogger.Instance, client);
            var check = await service.CheckLatestAsync("1.3", CancellationToken.None);

            var result = await service.DownloadAndVerifyAsync(check.LatestRelease, null, CancellationToken.None);

            Assert.False(result.Succeeded);
            Assert.Contains("SHA-256", result.Message);
        }

        [Fact]
        public async Task DownloadAndVerify_RejectsNonJsDelivrAssetUrls()
        {
            var client = CreateClient(_ => JsonResponse(FeedJson(
                "1.4",
                "https://example.test/AccuXSetup-1.4.exe",
                "https://example.test/AccuXSetup-1.4.exe.sha256")));
            var service = new JsDelivrReleaseService(NullLogger.Instance, client);
            var check = await service.CheckLatestAsync("1.3", CancellationToken.None);

            var result = await service.DownloadAndVerifyAsync(check.LatestRelease, null, CancellationToken.None);

            Assert.False(result.Succeeded);
            Assert.Contains("jsDelivr", result.Message);
        }

        private static HttpClient CreateClient(Func<HttpRequestMessage, HttpResponseMessage> responder)
        {
            return new HttpClient(new StubHandler(responder));
        }

        private static string FeedJson(
            string version,
            string installerUrl = null,
            string checksumUrl = null,
            string tag = null)
        {
            var installerName = "AccuXSetup-" + version + ".exe";
            var assetRoot = "https://cdn.jsdelivr.net/gh/smy116/AccuX@update-feed/releases/" + version;
            installerUrl = installerUrl ?? assetRoot + "/" + installerName;
            checksumUrl = checksumUrl ?? assetRoot + "/" + installerName + ".sha256";
            tag = tag ?? "v" + version;
            return "{\"version\":\"" + version + "\",\"tag\":\"" + tag + "\",\"name\":\"AccuX v" + version
                + "\",\"notes\":\"修复问题\",\"releaseNotesUrl\":\""
                + assetRoot + "/RELEASE-NOTES.md"
                + "\",\"installerUrl\":\"" + installerUrl + "\",\"sha256Url\":\"" + checksumUrl + "\"}";
        }

        private static HttpResponseMessage JsonResponse(string json)
        {
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(json, Encoding.UTF8, "application/json")
            };
        }

        private static HttpResponseMessage TextResponse(string text)
        {
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(text, Encoding.UTF8, "text/plain")
            };
        }

        private static string ComputeSha256(byte[] value)
        {
            using (var sha = SHA256.Create())
            {
                return BitConverter.ToString(sha.ComputeHash(value)).Replace("-", string.Empty).ToLowerInvariant();
            }
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
    }
}
