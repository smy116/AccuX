using System;
using System.IO;
using System.Net;
using System.Net.Http;
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

            using (var coordinator = CreateCoordinator(store, client))
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

            using (var coordinator = CreateCoordinator(store, client))
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

            using (var coordinator = CreateCoordinator(store, client))
            {
                coordinator.StartAutomaticCheck("1.3", null);
                await Task.Delay(100);
                Assert.Equal(0, requests);
            }
        }

        private static UpdateCoordinator CreateCoordinator(AccuXSettingsStore store, HttpClient client)
        {
            return new UpdateCoordinator(
                store,
                new JsDelivrReleaseService(NullLogger.Instance, client),
                new NoOpInstallerLauncher(),
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
            var installerUrl = JsDelivrReleaseService.GetAssetUrl(version, installerName);
            var checksumUrl = JsDelivrReleaseService.GetAssetUrl(version, installerName + ".sha256");
            var json = "{\"version\":\"" + version + "\",\"tag\":\"v" + version + "\","
                + "\"name\":\"AccuX\",\"notes\":\"\","
                + "\"releaseNotesUrl\":\"" + JsDelivrReleaseService.GetAssetUrl(version, "RELEASE-NOTES.md") + "\","
                + "\"installerUrl\":\"" + installerUrl + "\",\"sha256Url\":\"" + checksumUrl + "\"}";
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(json)
            };
        }

        private sealed class NoOpInstallerLauncher : IInstallerLauncher
        {
            public bool TryLaunch(string installerPath, out string errorMessage)
            {
                errorMessage = null;
                return true;
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
