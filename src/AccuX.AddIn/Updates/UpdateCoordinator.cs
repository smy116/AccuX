using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using AccuX.Core.Configuration;
using AccuX.Core.Logging;

namespace AccuX.AddIn.Updates
{
    /// <summary>
    /// 统一管理自动/手动检测、节流、下载和安装程序启动。
    /// </summary>
    internal sealed class UpdateCoordinator : IDisposable
    {
        private static readonly TimeSpan AutomaticCheckInterval = TimeSpan.FromHours(24);

        private readonly AccuXSettingsStore _settingsStore;
        private readonly JsDelivrReleaseService _releaseService;
        private readonly IInstallerLauncher _installerLauncher;
        private readonly ILogger _logger;
        private readonly Dispatcher _uiDispatcher;
        private readonly Func<DateTime> _utcNow;
        private readonly SemaphoreSlim _checkGate = new SemaphoreSlim(1, 1);
        private readonly CancellationTokenSource _cancellation = new CancellationTokenSource();
        private bool _disposed;

        public UpdateCoordinator(
            AccuXSettingsStore settingsStore,
            JsDelivrReleaseService releaseService,
            IInstallerLauncher installerLauncher,
            ILogger logger,
            Dispatcher uiDispatcher,
            Func<DateTime> utcNow = null)
        {
            _settingsStore = settingsStore ?? throw new ArgumentNullException(nameof(settingsStore));
            _releaseService = releaseService ?? throw new ArgumentNullException(nameof(releaseService));
            _installerLauncher = installerLauncher ?? throw new ArgumentNullException(nameof(installerLauncher));
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

        public async Task<UpdateInstallResult> DownloadAndLaunchAsync(
            UpdateCheckResult checkResult,
            IProgress<double> progress)
        {
            if (_disposed)
            {
                return UpdateInstallResult.Failed("AccuX 正在关闭，无法启动升级。");
            }

            if (checkResult == null || !checkResult.IsSuccessful || !checkResult.HasUpdate || checkResult.LatestRelease == null)
            {
                return UpdateInstallResult.Failed("没有可用的升级版本。");
            }

            var download = await _releaseService.DownloadAndVerifyAsync(
                checkResult.LatestRelease,
                progress,
                _cancellation.Token).ConfigureAwait(false);
            if (!download.Succeeded)
            {
                return download;
            }

            if (!_installerLauncher.TryLaunch(download.InstallerPath, out var errorMessage))
            {
                return UpdateInstallResult.Failed(errorMessage);
            }

            return download;
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

                if (!_releaseService.HasRequiredAssets(result.LatestRelease, out var assetError))
                {
                    _logger.Warn("自动升级检测发现新版本但附件不完整：" + assetError);
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
