using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using AccuX.Core.Configuration;
using AccuX.Core.Logging;

namespace AccuX.AddIn.Updates
{
    /// <summary>
    /// 统一管理自动/手动检测、节流和浏览器下载入口。
    /// </summary>
    internal sealed class UpdateCoordinator : IDisposable
    {
        private static readonly TimeSpan AutomaticCheckInterval = TimeSpan.FromHours(24);

        private readonly AccuXSettingsStore _settingsStore;
        private readonly GitHubReleaseService _releaseService;
        private readonly IBrowserLauncher _browserLauncher;
        private readonly ILogger _logger;
        private readonly Dispatcher _uiDispatcher;
        private readonly Func<DateTime> _utcNow;
        private readonly SemaphoreSlim _checkGate = new SemaphoreSlim(1, 1);
        private readonly CancellationTokenSource _cancellation = new CancellationTokenSource();
        private bool _disposed;

        public UpdateCoordinator(
            AccuXSettingsStore settingsStore,
            GitHubReleaseService releaseService,
            IBrowserLauncher browserLauncher,
            ILogger logger,
            Dispatcher uiDispatcher,
            Func<DateTime> utcNow = null)
        {
            _settingsStore = settingsStore ?? throw new ArgumentNullException(nameof(settingsStore));
            _releaseService = releaseService ?? throw new ArgumentNullException(nameof(releaseService));
            _browserLauncher = browserLauncher ?? throw new ArgumentNullException(nameof(browserLauncher));
            _logger = logger ?? NullLogger.Instance;
            _uiDispatcher = uiDispatcher ?? throw new ArgumentNullException(nameof(uiDispatcher));
            _utcNow = utcNow ?? (() => DateTime.UtcNow);
        }

        public void StartAutomaticCheck(string currentVersion, Action<UpdateCheckResult> onUpdateAvailable)
        {
            if (_disposed || !IsAutomaticCheckDue())
            {
                return;
            }

            _ = RunAutomaticCheckAsync(currentVersion, onUpdateAvailable);
        }

        public Task<UpdateCheckResult> CheckManuallyAsync(string currentVersion)
        {
            if (_disposed)
            {
                return Task.FromResult(UpdateCheckResult.Failed(currentVersion, "AccuX 正在关闭，无法检测升级。"));
            }

            return CheckAndRecordAsync(currentVersion, _cancellation.Token);
        }

        public UpdateOpenResult OpenInstallerDownload(UpdateCheckResult checkResult)
        {
            if (_disposed)
            {
                return UpdateOpenResult.Failed("AccuX 正在关闭，无法打开升级下载地址。");
            }

            if (checkResult == null || !checkResult.IsSuccessful || !checkResult.HasUpdate || checkResult.LatestRelease == null)
            {
                return UpdateOpenResult.Failed("没有可用的升级版本。");
            }

            if (!_releaseService.HasInstaller(checkResult.LatestRelease, out var assetError))
            {
                return UpdateOpenResult.Failed(assetError);
            }

            var url = checkResult.LatestRelease.InstallerUrl;
            if (!_browserLauncher.TryOpen(url, out var errorMessage))
            {
                return UpdateOpenResult.Failed(errorMessage);
            }

            return UpdateOpenResult.Success();
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            try { _cancellation.Cancel(); } catch { }
            _cancellation.Dispose();
            // 不在这里释放 SemaphoreSlim：停止时可能仍有一个等待中的检查在执行 finally。
        }

        private bool IsAutomaticCheckDue()
        {
            var settings = _settingsStore.Current;
            if (!settings.AutoCheckForUpdates)
            {
                return false;
            }

            if (!settings.LastUpdateCheckUtc.HasValue)
            {
                return true;
            }

            var elapsed = _utcNow() - settings.LastUpdateCheckUtc.Value.ToUniversalTime();
            return elapsed >= AutomaticCheckInterval;
        }

        private async Task RunAutomaticCheckAsync(
            string currentVersion,
            Action<UpdateCheckResult> onUpdateAvailable)
        {
            try
            {
                var result = await CheckAndRecordAsync(currentVersion, _cancellation.Token).ConfigureAwait(false);
                if (_disposed || result == null)
                {
                    return;
                }

                if (!result.IsSuccessful)
                {
                    _logger.Warn("自动升级检测未成功：" + result.ErrorMessage);
                    return;
                }

                if (!result.HasUpdate || onUpdateAvailable == null)
                {
                    return;
                }

                if (!_settingsStore.Current.AutoCheckForUpdates)
                {
                    return;
                }

                _ = _uiDispatcher.BeginInvoke(new Action(() =>
                {
                    if (!_disposed && _settingsStore.Current.AutoCheckForUpdates)
                    {
                        onUpdateAvailable(result);
                    }
                }));
            }
            catch (OperationCanceledException)
            {
                // 插件关闭时取消后台检查，不再显示任何提示。
            }
            catch (Exception ex)
            {
                _logger.Warn("自动升级检测失败：" + ex.Message);
            }
        }

        private async Task<UpdateCheckResult> CheckAndRecordAsync(
            string currentVersion,
            CancellationToken cancellationToken)
        {
            await _checkGate.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                var result = await _releaseService.CheckLatestAsync(currentVersion, cancellationToken)
                    .ConfigureAwait(false);
                _settingsStore.RecordUpdateCheck(_utcNow());
                return result;
            }
            finally
            {
                _checkGate.Release();
            }
        }
    }
}
