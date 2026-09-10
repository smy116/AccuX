using System;
using System.Reflection;
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
            var basicConfig = Config.GetSection("basicFinance", new BasicFinanceConfig());

            // Host 层：宿主差异与 COM 边界。
            var hostOptions = new HostOptions
            {
                LargeSelectionWarning = basicConfig.LargeSelectionWarning,
                MaxProcessCells = basicConfig.MaxProcessCells
            };

            // Host 层：宿主差异与 COM 边界。一个 Core 窄接口对应一个实现类（规格 §5.2）。
            var rangeHost = new ExcelRangeOperationHost(_application, hostOptions);
            var directoryHost = new ExcelWorkbookDirectoryHost(_application);
            var commentHost = new ExcelCellCommentHost(_application);
            var markHost = new ExcelCellMarkHost(_application);
            var pipeline = new RangeOperationPipeline(rangeHost, Logger);

            var context = new ModuleContext(
                Config,
                Logger,
                rangeHost.Context,
                pipeline,
                AccuXVersion,
                directoryHost,
                commentHost,
                markHost);
            Dispatcher = new CommandDispatcher(context, Logger);

            // 用户交互统一由 WPF 实现；模块通过 IUserPrompt 使用。
            Prompt = new WpfUserPrompt(rangeHost.Context, basicConfig.LargeSelectionWarning, basicConfig.RoundDigits);

            // 显式模块注册（V1 不做目录扫描 / 反射发现 / 热加载）。
            _moduleRegistry = new ModuleRegistry(Logger);
            var basicFinance = new BasicFinanceModule { PromptOverride = Prompt };
            _moduleRegistry.Register(basicFinance, context, Dispatcher);

            var mark = new MarkModule();
            _moduleRegistry.Register(mark, context, Dispatcher);

            Logger.Info("AccuX 启动完成，已注册命令数：" + Dispatcher.Commands.Count);
        }

        /// <summary>
        /// 停止 AccuX：卸载模块。
        /// </summary>
        public void Stop()
        {
            try
            {
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

                return assembly.GetName().Version?.ToString() ?? "1.0.0";
            }
            catch
            {
                return "1.0.0";
            }
        }
    }
}
