using System;
using System.Collections.Generic;
using AccuX.Core.Commands;
using AccuX.Core.Configuration;
using AccuX.Core.Modules;
using AccuX.Core.Operations;
using AccuX.Modules.BasicFinance.AmountConversion;
using AccuX.Modules.BasicFinance.ChineseAmount;
using AccuX.Modules.BasicFinance.Comment;
using AccuX.Modules.BasicFinance.Common;
using AccuX.Modules.BasicFinance.Directory;
using AccuX.Modules.BasicFinance.Rounding;
using AccuX.Modules.BasicFinance.SelectionSum;
using AccuX.Modules.BasicFinance.UI;

namespace AccuX.Modules.BasicFinance
{
    /// <summary>
    /// AccuX 唯一业务模块（规格 §5.4 / §6）。
    /// 负责基础功能的 Ribbon Command 定义、用户输入校验与业务算法调用；
    /// 不解决 Excel/WPS 差异，不直接持有 COM 对象。
    /// </summary>
    public sealed class BasicFinanceModule : IAccuXModule
    {
        public const string ModuleId = "AccuX.Modules.BasicFinance";

        private IAccuXContext _context;
        private IUserPrompt _prompt;
        private ICommentPrompt _commentPrompt;
        private CommandDefinition[] _commands = Array.Empty<CommandDefinition>();

        public string Id
        {
            get { return ModuleId; }
        }

        /// <summary>
        /// 允许外部注入自定义 IUserPrompt（单元测试用），默认使用 WPF 实现。
        /// </summary>
        public IUserPrompt PromptOverride { get; set; }

        /// <summary>
        /// 允许测试或其他宿主替换批注窗口交互；未设置时使用默认 WPF 实现。
        /// </summary>
        public ICommentPrompt CommentPromptOverride { get; set; }

        public void Initialize(IAccuXContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));

            var config = context.Config.GetSection("basicFinance", new BasicFinanceConfig());
            var settings = context.Config.GetSection("settings", AccuXSettings.CreateDefault());
            var defaultPrompt = new WpfUserPrompt(context.Host, settings.LargeSelectionWarning, config.RoundDigits);
            _prompt = PromptOverride ?? defaultPrompt;
            _commentPrompt = CommentPromptOverride
                ?? PromptOverride as ICommentPrompt
                ?? defaultPrompt;

            var stateOptions = new HostStateOptions
            {
                DisableScreenUpdating = true,
                DisableEvents = true,
                DisableDisplayAlerts = true,
                ManualCalculation = false
            };

            _commands = new[]
            {
                new RoundingCommand(_prompt, stateOptions).CreateDefinition(ModuleId),
                new AmountConversionCommand(_prompt, stateOptions).CreateDefinition(ModuleId),
                new SelectionSumCommand(_prompt).CreateDefinition(ModuleId),
                new ChineseAmountCommand(_prompt, stateOptions).CreateDefinition(ModuleId),
                new DirectoryCommand(_prompt).CreateDefinition(ModuleId),
                new CommentCommand(_commentPrompt).CreateDefinition(ModuleId)
            };

            context.Logger.Info("BasicFinance 模块初始化完成。");
        }

        public IEnumerable<CommandDefinition> GetCommands()
        {
            return _commands;
        }

        public void Shutdown()
        {
            _commands = Array.Empty<CommandDefinition>();
        }
    }
}
