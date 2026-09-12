using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using AccuX.AddIn.Updates;
using AccuX.Core.Configuration;
using AccuX.Core.Logging;
using Xunit;

namespace AccuX.AddIn.Tests
{
    public sealed class UpdateCoordinatorTests : IDisposable
    {
        private readonly string _path = Path.Combine(Path.GetTempPath(), "accux-update-coordinator-" + Guid.NewGuid() + ".json");

        public void Dispose()
        {
            try
            {
                if (File.Exists(_path)) File.Delete(_path);
            }
            catch
            {
                // 测试清理失败不影响断言结果。
            }
        }

        [Fact]
        public async Task ManualCheck_BypassesAutomaticThrottle()
        {
            var requests = 0;
            var client = CreateClient(_ =>
            {
                Interlocked.Increment(ref requests);
                return LatestResponse("1.4");
            });
            var store = new AccuXSettingsStore(new JsonConfigManager(_path));
            store.RecordUpdateCheck(DateTime.UtcNow);

            using (var coordinator = CreateCoordinator(store, client, new RecordingBrowserLauncher()))
            {
                var first = await coordinator.CheckManuallyAsync("1.3");
                var second = await coordinator.CheckManuallyAsync("1.3");

                Assert.True(first.IsSuccessful);
                Assert.True(second.IsSuccessful);
                Assert.Equal(2, requests);
            }
        }

        [Fact]
        public async Task AutomaticCheck_IsLimitedToOncePerDay()
        {
            var requests = 0;
            var completed = new TaskCompletionSource<bool>();
            var client = CreateClient(_ =>
            {
                Interlocked.Increment(ref requests);
                completed.TrySetResult(true);
                return LatestResponse("1.3");
            });
            var store = new AccuXSettingsStore(new JsonConfigManager(_path));

            using (var coordinator = CreateCoordinator(store, client, new RecordingBrowserLauncher()))
            {
                coordinator.StartAutomaticCheck("1.3", null);
                var finished = await Task.WhenAny(completed.Task, Task.Delay(TimeSpan.FromSeconds(5)));
                Assert.Same(completed.Task, finished);

                var deadline = DateTime.UtcNow.AddSeconds(5);
                while (!store.Current.LastUpdateCheckUtc.HasValue && DateTime.UtcNow < deadline)
                {
                    await Task.Delay(20);
                }

                Assert.True(store.Current.LastUpdateCheckUtc.HasValue);
                coordinator.StartAutomaticCheck("1.3", null);
                await Task.Delay(100);
                Assert.Equal(1, requests);
            }
        }

        [Fact]
        public async Task AutomaticCheck_DoesNothingWhenDisabled()
        {
            var requests = 0;
            var client = CreateClient(_ =>
            {
                Interlocked.Increment(ref requests);
                return LatestResponse("1.4");
            });
            var store = new AccuXSettingsStore(new JsonConfigManager(_path));
            store.SaveEditable(new AccuXSettings { AutoCheckForUpdates = false });

            using (var coordinator = CreateCoordinator(store, client, new RecordingBrowserLauncher()))
            {
                coordinator.StartAutomaticCheck("1.3", null);
                await Task.Delay(100);
                Assert.Equal(0, requests);
            }
        }

        [Fact]
        public async Task OpenInstallerDownload_UsesBrowserAssetUrl()
        {
            var browser = new RecordingBrowserLauncher();
            var client = CreateClient(_ => LatestResponse("1.4"));
            var store = new AccuXSettingsStore(new JsonConfigManager(_path));

            using (var coordinator = CreateCoordinator(store, client, browser))
            {
                var check = await coordinator.CheckManuallyAsync("1.3");
                var result = coordinator.OpenInstallerDownload(check);

                Assert.True(result.Succeeded);
                Assert.Equal("https://github.com/smy116/AccuX/releases/download/v1.4/AccuXSetup-1.4.exe", browser.Url);
            }
        }

        [Fact]
        public async Task OpenInstallerDownload_ReportsBrowserFailure()
        {
            var browser = new RecordingBrowserLauncher { ShouldFail = true };
            var store = new AccuXSettingsStore(new JsonConfigManager(_path));

            using (var coordinator = CreateCoordinator(store, CreateClient(_ => LatestResponse("1.4")), browser))
            {
                var check = await coordinator.CheckManuallyAsync("1.3");
                var result = coordinator.OpenInstallerDownload(check);

                Assert.False(result.Succeeded);
                Assert.Contains("浏览器启动失败", result.Message);
            }
        }

        [Fact]
        public async Task OpenInstallerDownload_RejectsWhenAlreadyUpToDate()
        {
            var browser = new RecordingBrowserLauncher();
            var store = new AccuXSettingsStore(new JsonConfigManager(_path));

            using (var coordinator = CreateCoordinator(store, CreateClient(_ => LatestResponse("1.3")), browser))
            {
                var check = await coordinator.CheckManuallyAsync("1.3");
                var result = coordinator.OpenInstallerDownload(check);

                Assert.False(result.Succeeded);
                Assert.Contains("没有可用", result.Message);
                Assert.Null(browser.Url);
            }
        }

        [Fact]
        public void OpenInstallerDownload_RejectsAfterDispose()
        {
            var browser = new RecordingBrowserLauncher();
            var store = new AccuXSettingsStore(new JsonConfigManager(_path));
            var coordinator = CreateCoordinator(store, CreateClient(_ => LatestResponse("1.4")), browser);

            coordinator.Dispose();
            var result = coordinator.OpenInstallerDownload(null);

            Assert.False(result.Succeeded);
            Assert.Contains("关闭", result.Message);
        }

        private static UpdateCoordinator CreateCoordinator(
            AccuXSettingsStore store,
            HttpClient client,
            IBrowserLauncher browser)
        {
            return new UpdateCoordinator(
                store,
                new GitHubReleaseService(NullLogger.Instance, client),
                browser,
                NullLogger.Instance,
                Dispatcher.CurrentDispatcher);
        }

        private static HttpClient CreateClient(Func<HttpRequestMessage, HttpResponseMessage> responder)
        {
            return new HttpClient(new StubHandler(responder));
        }

        private static HttpResponseMessage LatestResponse(string version)
        {
            var installerName = "AccuXSetup-" + version + ".exe";
            var tag = "v" + version;
            var installerUrl = "https://github.com/smy116/AccuX/releases/download/" + tag + "/" + installerName;
            var json = "{\"tag_name\":\"" + tag + "\",\"name\":\"AccuX\",\"body\":\"\","
                + "\"html_url\":\"https://github.com/smy116/AccuX/releases/tag/" + tag + "\","
                + "\"published_at\":\"2026-01-01T00:00:00Z\",\"draft\":false,\"prerelease\":false,"
                + "\"assets\":[{\"name\":\"" + installerName + "\",\"browser_download_url\":\"" + installerUrl + "\"}]}";
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(json, Encoding.UTF8, "application/json")
            };
        }

        private sealed class RecordingBrowserLauncher : IBrowserLauncher
        {
            public bool ShouldFail { get; set; }

            public string Url { get; private set; }

            public bool TryOpen(string url, out string errorMessage)
            {
                Url = url;
                errorMessage = ShouldFail ? "浏览器启动失败" : null;
                return !ShouldFail;
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
