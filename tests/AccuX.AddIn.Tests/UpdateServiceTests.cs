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
        public async Task CheckLatest_ReturnsNewStableRelease()
        {
            var client = CreateClient(request => JsonResponse(
                "{\"tag_name\":\"v1.4\",\"name\":\"AccuX v1.4\",\"body\":\"修复问题\","
                + "\"html_url\":\"https://github.com/smy116/AccuX/releases/tag/v1.4\","
                + "\"draft\":false,\"prerelease\":false,\"assets\":[]}"));
            var service = new GitHubReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.True(result.IsSuccessful);
            Assert.True(result.HasUpdate);
            Assert.Equal("1.4", result.LatestRelease.Version.Text);
        }

        [Fact]
        public async Task CheckLatest_HttpErrorReturnsFailure()
        {
            var client = CreateClient(_ => new HttpResponseMessage(HttpStatusCode.NotFound));
            var service = new GitHubReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.False(result.IsSuccessful);
            Assert.Contains("HTTP 404", result.ErrorMessage);
        }

        [Theory]
        [InlineData("1.4", "v1.4", false)]
        [InlineData("1.10", "v1.4", false)]
        [InlineData("1.4", "v1.10", true)]
        public async Task CheckLatest_ComparesVersionsNumerically(string currentVersion, string tag, bool expectedUpdate)
        {
            var client = CreateClient(_ => JsonResponse(ReleaseJson(tag, false)));
            var service = new GitHubReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync(currentVersion, CancellationToken.None);

            Assert.True(result.IsSuccessful);
            Assert.Equal(expectedUpdate, result.HasUpdate);
        }

        [Theory]
        [InlineData("v1.4-beta", true)]
        [InlineData("v1.4.1", false)]
        public async Task CheckLatest_RejectsPrereleaseOrUnsupportedTag(string tag, bool prerelease)
        {
            var client = CreateClient(_ => JsonResponse(ReleaseJson(tag, prerelease)));
            var service = new GitHubReleaseService(NullLogger.Instance, client);

            var result = await service.CheckLatestAsync("1.3", CancellationToken.None);

            Assert.False(result.IsSuccessful);
        }

        [Fact]
        public async Task DownloadAndVerify_RequiresMatchingReleaseChecksum()
        {
            var installerName = "AccuXSetup-1.4.exe";
            var installerBytes = Encoding.UTF8.GetBytes("test installer");
            var hash = ComputeSha256(installerBytes);
            var responses = new Dictionary<string, HttpResponseMessage>(StringComparer.Ordinal)
            {
                [LatestReleaseUrl] = LatestReleaseResponse(),
                ["https://example.test/checksum"] = TextResponse(hash + " *" + installerName),
                ["https://example.test/installer"] = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ByteArrayContent(installerBytes)
                }
            };
            var client = CreateClient(request => responses[request.RequestUri.AbsoluteUri]);
            var service = new GitHubReleaseService(NullLogger.Instance, client);
            var check = await service.CheckLatestAsync("1.3", CancellationToken.None);
            check.LatestRelease.Assets = new[]
            {
                new GitHubReleaseAsset
                {
                    Name = installerName,
                    BrowserDownloadUrl = "https://example.test/installer"
                },
                new GitHubReleaseAsset
                {
                    Name = installerName + ".sha256",
                    BrowserDownloadUrl = "https://example.test/checksum"
                }
            };

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
            var responses = new Dictionary<string, HttpResponseMessage>(StringComparer.Ordinal)
            {
                [LatestReleaseUrl] = LatestReleaseResponse(),
                ["https://example.test/checksum"] = TextResponse(new string('0', 64) + " *" + installerName),
                ["https://example.test/installer"] = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ByteArrayContent(Encoding.UTF8.GetBytes("test installer"))
                }
            };
            var client = CreateClient(request => responses[request.RequestUri.AbsoluteUri]);
            var service = new GitHubReleaseService(NullLogger.Instance, client);
            var check = await service.CheckLatestAsync("1.3", CancellationToken.None);
            check.LatestRelease.Assets = new[]
            {
                new GitHubReleaseAsset { Name = installerName, BrowserDownloadUrl = "https://example.test/installer" },
                new GitHubReleaseAsset { Name = installerName + ".sha256", BrowserDownloadUrl = "https://example.test/checksum" }
            };

            var result = await service.DownloadAndVerifyAsync(check.LatestRelease, null, CancellationToken.None);

            Assert.False(result.Succeeded);
            Assert.Contains("SHA-256", result.Message);
        }

        private static HttpClient CreateClient(Func<HttpRequestMessage, HttpResponseMessage> responder)
        {
            return new HttpClient(new StubHandler(responder));
        }

        private static string LatestReleaseUrl => "https://api.github.com/repos/smy116/AccuX/releases/latest";

        private static HttpResponseMessage LatestReleaseResponse()
        {
            return JsonResponse(ReleaseJson("v1.4", false));
        }

        private static string ReleaseJson(string tag, bool prerelease)
        {
            return "{\"tag_name\":\"" + tag + "\",\"name\":\"AccuX\",\"body\":\"\","
                + "\"html_url\":\"https://github.com/smy116/AccuX/releases\","
                + "\"draft\":false,\"prerelease\":" + prerelease.ToString().ToLowerInvariant() + ",\"assets\":[]}";
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
