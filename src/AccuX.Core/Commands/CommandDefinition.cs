using System;
using System.Collections.Generic;

namespace AccuX.Core.Commands
{
    /// <summary>
    /// 命令执行委托。参数为命令 ID，返回值决定是否成功（用于 Ribbon 状态等场景）。
    /// </summary>
    public delegate CommandResult CommandHandler(CommandExecutionContext executionContext);

    /// <summary>
    /// 命令定义（规格 §6 / §7）。
    /// Ribbon 按钮 ID 映射到 Command ID；模块只注册 Command，不动态贡献 Ribbon UI。
    /// </summary>
    public sealed class CommandDefinition
    {
        public CommandDefinition(
            string id,
            string displayName,
            string moduleId,
            CommandHandler handler,
            string description = null,
            string ribbonControlId = null)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                throw new ArgumentException("命令 ID 不能为空。", nameof(id));
            }

            Id = id;
            DisplayName = displayName ?? id;
            ModuleId = moduleId ?? string.Empty;
            Handler = handler ?? throw new ArgumentNullException(nameof(handler));
            Description = description ?? string.Empty;
            RibbonControlId = ribbonControlId ?? id;
        }

        /// <summary>命令 ID，例如 accux.basic.round。</summary>
        public string Id { get; }

        public string DisplayName { get; }

        public string ModuleId { get; }

        public string Description { get; }

        /// <summary>对应的 Ribbon 控件 ID（V1 与命令 ID 一致）。</summary>
        public string RibbonControlId { get; }

        public CommandHandler Handler { get; }
    }

    /// <summary>
    /// 命令执行上下文，由 CommandDispatcher 构造。
    /// </summary>
    public sealed class CommandExecutionContext
    {
        public CommandExecutionContext(CommandDefinition definition, Modules.IAccuXContext context)
        {
            Definition = definition ?? throw new ArgumentNullException(nameof(definition));
            Context = context ?? throw new ArgumentNullException(nameof(context));
        }

        public CommandDefinition Definition { get; }

        public Modules.IAccuXContext Context { get; }
    }
}
