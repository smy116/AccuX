using System;
using System.Reflection;
using System.Threading;
using System.Windows;
using WpfDispatcher = System.Windows.Threading.Dispatcher;
using AccuX.AddIn.Settings;
using AccuX.AddIn.Updates;
using AccuX.Core.Commands;
using AccuX.Core.Configuration;
using AccuX.Core.Logging;
using AccuX.Core.Modules;
using AccuX.Core.Operations;
using AccuX.Host;
using AccuX.Host.Features;
using AccuX.Modules.BasicFinance;
using AccuX.Modules.BasicFinance.Common;
using AccuX.Modules.BasicFinance.UI;
using AccuX.Modules.Mark;
using AccuX.Modules.Compare;
using Excel = Microsoft.Office.Interop.Excel;

namespace AccuX.AddIn
{
    /// <summary>
    /// Add-in 组合根（规格 §5.1 / §4）。
    /// <para>
    /// 创建并注入具体实现：Config → Logger → Host → Pipeline → 模块显式注册。
    /// 依赖方向固定为 AddIn → Host + Modules + Core，Core 不引用 Host。
    /// </para>
    /// </summary>
    internal sealed class AddInCompositionRoot
    {
        private readonly Excel.Application _application;
        private ModuleRegistry _moduleRegistry;
        private AccuXSettingsStore _settingsStore;
        private AccuXSettings _settings;
        private HostOptions _hostOptions;
        private ExcelRangeOperationHost _rangeHost;
        private ExcelRegionCompareHost _compareHost;
        private WpfUserPrompt _wpfPrompt;
        private UpdateCoordinator _updateCoordinator;
        private SettingsWindow _settingsWindow;
        private WpfDispatcher _uiDispatcher;
        private bool _started;

        public AddInCompositionRoot(Excel.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public ILogger Logger { get; private set; }

        public IConfigManager Config { get; private set; }

        public CommandDispatcher Dispatcher { get; private set; }

        public IUserPrompt Prompt { get; private set; }

        public string AccuXVersion { get; private set; }

        /// <summary>
        /// 启动 AccuX：初始化基础设施并显式注册 V1 模块。
        /// </summary>
        public void Start()
        {
            AccuXVersion = ResolveVersion();

            Logger = new FileLogger(FileLogger.DefaultDirectory, LogLevel.Info);
            Logger.Info("AccuX 启动，版本 " + AccuXVersion);

            Config = new JsonConfigManager(JsonConfigManager.DefaultConfigPath);
            _settingsStore = new AccuXSettingsStore(Config, Logger);
            _settings = _settingsStore.Load();
            var basicConfig = Config.GetSection("basicFinance", new BasicFinanceConfig());

            // Host 层：宿主差异与 COM 边界。
            _hostOptions = new HostOptions
            {
                LargeSelectionWarning = _settings.LargeSelectionWarning,
                MaxProcessCells = _settings.MaxProcessCells
            };

            // Host 层：宿主差异与 COM 边界。一个 Core 窄接口对应一个实现类（规格 §5.2）。
            _rangeHost = new ExcelRangeOperationHost(_application, _hostOptions);
            var directoryHost = new ExcelWorkbookDirectoryHost(_application);
            var commentHost = new ExcelCellCommentHost(_application);
            var markHost = new ExcelCellMarkHost(_application);
            _compareHost = new ExcelRegionCompareHost(_application, _hostOptions);
            var pipeline = new RangeOperationPipeline(_rangeHost, Logger);

            var context = new ModuleContext(
                Config,
                Logger,
                _rangeHost.Context,
                pipeline,
                AccuXVersion,
                directoryHost,
                commentHost,
                markHost,
                _compareHost);
            Dispatcher = new CommandDispatcher(context, Logger);

            // 用户交互统一由 WPF 实现；模块通过 IUserPrompt 使用。
            _wpfPrompt = new WpfUserPrompt(_rangeHost.Context, _settings.LargeSelectionWarning, basicConfig.RoundDigits);
            Prompt = _wpfPrompt;

            // 显式模块注册（V1 不做目录扫描 / 反射发现 / 热加载）。
            _moduleRegistry = new ModuleRegistry(Logger);
            var basicFinance = new BasicFinanceModule { PromptOverride = Prompt };
            _moduleRegistry.Register(basicFinance, context, Dispatcher);

            var mark = new MarkModule();
            _moduleRegistry.Register(mark, context, Dispatcher);

            var compare = new RegionCompareModule();
            _moduleRegistry.Register(compare, context, Dispatcher);

            Dispatcher.Register(new CommandDefinition(
                "accux.settings.open",
                "设置",
                "AccuX.AddIn.Settings",
                OpenSettings,
                "查看 AccuX 版本、处理限制和升级设置。",
                "accux.settings.open"));

            _uiDispatcher = WpfDispatcher.FromThread(Thread.CurrentThread) ?? WpfDispatcher.CurrentDispatcher;
            _updateCoordinator = new UpdateCoordinator(
                _settingsStore,
                new GitHubReleaseService(Logger),
                new BrowserLauncher(),
                Logger,
                _uiDispatcher);
            _started = true;
            _updateCoordinator.StartAutomaticCheck(AccuXVersion, ShowAutomaticUpdate);

            Logger.Info("AccuX 启动完成，已注册命令数：" + Dispatcher.Commands.Count);
        }

        /// <summary>
        /// 停止 AccuX：卸载模块。
        /// </summary>
        public void Stop()
        {
            try
            {
                try { _settingsWindow?.Close(); } catch { }
                _settingsWindow = null;
                _updateCoordinator?.Dispose();
                _updateCoordinator = null;
                _started = false;
                _moduleRegistry?.Dispose();
                _moduleRegistry = null;
                Logger?.Info("AccuX 已停止。");
            }
            catch
            {
                // 停止阶段忽略异常。
            }
        }

        /// <summary>
        /// 分发 Ribbon 命令。
        /// </summary>
        public CommandResult Dispatch(string commandId)
        {
            if (Dispatcher == null)
            {
                return CommandResult.Failed("AccuX 尚未初始化完成，请重新启动 Excel。");
            }

            return Dispatcher.Execute(commandId);
        }

        private CommandResult OpenSettings(CommandExecutionContext execution)
        {
            ShowSettings(null);
            return CommandResult.Cancelled();
        }

        private void ShowSettings(UpdateCheckResult initialCheckResult)
        {
            if (!_started || _settingsStore == null || _updateCoordinator == null)
            {
                return;
            }

            if (_settingsWindow != null)
            {
                if (initialCheckResult != null)
                {
                    _settingsWindow.ShowCheckResult(initialCheckResult, false);
                }

                try { _settingsWindow.Activate(); } catch { }
                return;
            }

            var window = new SettingsWindow(
                _settingsStore.Current,
                AccuXVersion,
                _rangeHost?.Context?.MainWindowHandle ?? IntPtr.Zero,
                _settingsStore,
                _updateCoordinator,
                ApplySettings,
                initialCheckResult);
            _settingsWindow = window;
            try
            {
                window.ShowDialog();
            }
            finally
            {
                if (ReferenceEquals(_settingsWindow, window))
                {
                    _settingsWindow = null;
                }
            }
        }

        private void ApplySettings(AccuXSettings settings)
        {
            var saved = _settingsStore.SaveEditable(settings);
            _settings = saved;
            if (_hostOptions != null)
            {
                _hostOptions.LargeSelectionWarning = saved.LargeSelectionWarning;
                _hostOptions.MaxProcessCells = saved.MaxProcessCells;
            }

            _wpfPrompt?.UpdateLargeSelectionWarning(saved.LargeSelectionWarning);
            Logger?.Info("应用设置已更新：警告阈值=" + saved.LargeSelectionWarning
                + "，最大单元格数=" + saved.MaxProcessCells
                + "，自动检测升级=" + saved.AutoCheckForUpdates);
        }

        private void ShowAutomaticUpdate(UpdateCheckResult result)
        {
            if (!_started || result?.LatestRelease?.Version == null)
            {
                return;
            }

            if (_settingsWindow != null)
            {
                _settingsWindow.ShowCheckResult(result, false);
                try { _settingsWindow.Activate(); } catch { }
                return;
            }

            var version = result.LatestRelease.Version.Text;
            var choice = MessageBox.Show(
                "发现 AccuX 新版本 " + version + "，是否打开设置查看并升级？",
                "AccuX 升级",
                MessageBoxButton.YesNo,
                MessageBoxImage.Information);
            if (choice == MessageBoxResult.Yes)
            {
                ShowSettings(result);
            }
        }

        private static string ResolveVersion()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var attribute = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
                if (attribute != null && !string.IsNullOrWhiteSpace(attribute.InformationalVersion))
                {
                    return attribute.InformationalVersion;
                }

                return assembly.GetName().Version?.ToString() ?? "1.4.0";
            }
            catch
            {
                return "1.4.0";
            }
        }
    }
}
