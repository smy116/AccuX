using System;
using System.Diagnostics;
using System.Globalization;
using System.Windows;
using System.Windows.Interop;
using AccuX.AddIn.Updates;
using AccuX.Core.Configuration;

namespace AccuX.AddIn.Settings
{
    /// <summary>
    /// AccuX 应用级设置窗口。布局由 XAML 定义，配置保存与升级流程由组合根提供。
    /// </summary>
    internal sealed partial class SettingsWindow : Window
    {
        private readonly AccuXSettings _initialSettings;
        private readonly string _currentVersion;
        private readonly AccuXSettingsStore _settingsStore;
        private readonly UpdateCoordinator _updateCoordinator;
        private readonly Action<AccuXSettings> _applySettings;

        private UpdateCheckResult _lastCheckResult;
        private bool _busy;

        public SettingsWindow(
            AccuXSettings settings,
            string currentVersion,
            IntPtr ownerHandle,
            AccuXSettingsStore settingsStore,
            UpdateCoordinator updateCoordinator,
            Action<AccuXSettings> applySettings,
            UpdateCheckResult initialCheckResult = null)
        {
            _initialSettings = (settings ?? AccuXSettings.CreateDefault()).Clone();
            _currentVersion = currentVersion ?? string.Empty;
            _settingsStore = settingsStore ?? throw new ArgumentNullException(nameof(settingsStore));
            _updateCoordinator = updateCoordinator ?? throw new ArgumentNullException(nameof(updateCoordinator));
            _applySettings = applySettings ?? throw new ArgumentNullException(nameof(applySettings));

            InitializeComponent();

            _currentVersionText.Text = "当前版本：" + FormatVersion(_currentVersion);
            _configPathText.Text = "配置文件：" + _settingsStore.ConfigFilePath;
            _warningBox.Text = _initialSettings.LargeSelectionWarning.ToString(CultureInfo.InvariantCulture);
            _maxBox.Text = _initialSettings.MaxProcessCells.ToString(CultureInfo.InvariantCulture);
            _autoCheckBox.IsChecked = _initialSettings.AutoCheckForUpdates;

            if (ownerHandle != IntPtr.Zero)
            {
                new WindowInteropHelper(this).Owner = ownerHandle;
            }

            ShowCheckResult(initialCheckResult, false);
            UpdateLastCheckText(_settingsStore.Current.LastUpdateCheckUtc);
        }

        public void ShowCheckResult(UpdateCheckResult result, bool showFailureMessage)
        {
            if (result == null)
            {
                _updateStatusText.Text = "尚未检测升级。";
                return;
            }

            _lastCheckResult = result;
            UpdateLastCheckText(_settingsStore.Current.LastUpdateCheckUtc);
            _releaseNotesBox.Visibility = Visibility.Collapsed;
            _installButton.IsEnabled = false;
            _releasePageButton.IsEnabled = false;

            if (!result.IsSuccessful)
            {
                _updateStatusText.Text = result.ErrorMessage;
                if (showFailureMessage)
                {
                    MessageBox.Show(this, result.ErrorMessage, "AccuX", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                return;
            }

            var release = result.LatestRelease;
            if (release == null)
            {
                _updateStatusText.Text = "未找到可用的稳定版本。";
                return;
            }

            _releasePageButton.IsEnabled = !string.IsNullOrWhiteSpace(release.HtmlUrl);
            if (result.HasUpdate)
            {
                _updateStatusText.Text = "发现新版本：" + release.Version.Text + "（当前 " + FormatVersion(_currentVersion) + "）。";
                _installButton.IsEnabled = true;
                var notes = release.Body ?? string.Empty;
                if (notes.Length > 4000) notes = notes.Substring(0, 4000) + Environment.NewLine + "……";
                _releaseNotesBox.Text = string.IsNullOrWhiteSpace(notes) ? "该 Release 没有提供更新说明。" : notes;
                _releaseNotesBox.Visibility = Visibility.Visible;
            }
            else
            {
                _updateStatusText.Text = "当前已是最新版本（" + FormatVersion(_currentVersion) + "）。";
            }
        }

        private async void OnManualCheck(object sender, RoutedEventArgs e)
        {
            if (_busy)
            {
                return;
            }

            SetBusy(true, false);
            try
            {
                var result = await _updateCoordinator.CheckManuallyAsync(_currentVersion);
                ShowCheckResult(result, true);
            }
            catch (OperationCanceledException)
            {
                _updateStatusText.Text = "升级检测已取消。";
            }
            catch (Exception ex)
            {
                _updateStatusText.Text = "升级检测失败：" + ex.Message;
                MessageBox.Show(this, _updateStatusText.Text, "AccuX", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            finally
            {
                SetBusy(false, false);
            }
        }

        private async void OnInstall(object sender, RoutedEventArgs e)
        {
            if (_busy || _lastCheckResult == null || !_lastCheckResult.HasUpdate)
            {
                return;
            }

            var confirm = MessageBox.Show(
                this,
                "升级安装程序需要替换 AccuX 文件。请先保存工作簿并关闭 Excel/WPS，再继续启动安装程序。是否继续？",
                "AccuX 升级",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.Yes)
            {
                return;
            }

            SetBusy(true, true);
            try
            {
                var progress = new Progress<double>(value => _downloadProgress.Value = value);
                var result = await _updateCoordinator.DownloadAndLaunchAsync(_lastCheckResult, progress);
                if (result.Succeeded)
                {
                    MessageBox.Show(this, result.Message + Environment.NewLine + "请关闭 Excel/WPS 后完成安装。", "AccuX", MessageBoxButton.OK, MessageBoxImage.Information);
                    Close();
                }
                else
                {
                    _updateStatusText.Text = result.Message;
                    MessageBox.Show(this, result.Message, "AccuX", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (OperationCanceledException)
            {
                _updateStatusText.Text = "升级下载已取消。";
            }
            catch (Exception ex)
            {
                _updateStatusText.Text = "升级失败：" + ex.Message;
                MessageBox.Show(this, _updateStatusText.Text, "AccuX", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            finally
            {
                SetBusy(false, false);
            }
        }

        private void OnOpenReleasePage(object sender, RoutedEventArgs e)
        {
            var url = _lastCheckResult?.LatestRelease?.HtmlUrl;
            if (string.IsNullOrWhiteSpace(url))
            {
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo { FileName = url, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "无法打开更新说明：" + ex.Message, "AccuX", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void OnSave(object sender, RoutedEventArgs e)
        {
            if (!TryReadSettings(out var settings, out var error))
            {
                MessageBox.Show(this, error, "AccuX 设置", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                _applySettings(settings);
                DialogResult = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "设置保存失败：" + ex.Message, "AccuX 设置", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private bool TryReadSettings(out AccuXSettings settings, out string error)
        {
            settings = _initialSettings.Clone();
            error = null;
            if (!long.TryParse(_warningBox.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var warning)
                || warning <= 0)
            {
                error = "警告单元格数量必须是大于 0 的整数。";
                return false;
            }

            if (!long.TryParse(_maxBox.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var maximum)
                || maximum <= 0)
            {
                error = "最大单元格数量必须是大于 0 的整数。";
                return false;
            }

            settings.LargeSelectionWarning = warning;
            settings.MaxProcessCells = maximum;
            settings.AutoCheckForUpdates = _autoCheckBox.IsChecked == true;
            if (!settings.TryValidate(out error))
            {
                return false;
            }

            return true;
        }

        private void SetBusy(bool busy, bool downloading)
        {
            _busy = busy;
            _manualCheckButton.IsEnabled = !busy;
            _installButton.IsEnabled = !busy && _lastCheckResult != null && _lastCheckResult.HasUpdate;
            _releasePageButton.IsEnabled = !busy && _lastCheckResult?.LatestRelease != null;
            _downloadProgress.Visibility = downloading ? Visibility.Visible : Visibility.Collapsed;
            if (!downloading)
            {
                _downloadProgress.Value = 0;
            }
        }

        private void UpdateLastCheckText(DateTime? utc)
        {
            _lastCheckText.Text = utc.HasValue
                ? "上次检测：" + utc.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.CurrentCulture)
                : "上次检测：尚未检测";
        }

        private static string FormatVersion(string version)
        {
            if (string.IsNullOrWhiteSpace(version))
            {
                return "未知";
            }

            var text = version.Trim().TrimStart('v', 'V');
            var plus = text.IndexOf('+');
            if (plus >= 0) text = text.Substring(0, plus);
            var dash = text.IndexOf('-');
            if (dash >= 0) text = text.Substring(0, dash);
            var parts = text.Split('.');
            return parts.Length >= 2 ? parts[0] + "." + parts[1] : text;
        }
    }
}
