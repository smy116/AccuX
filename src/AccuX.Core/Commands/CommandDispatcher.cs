using System;
using System.Collections.Generic;
using AccuX.Core.Logging;
using AccuX.Core.Modules;

namespace AccuX.Core.Commands
{
    /// <summary>
    /// 统一命令分发器（规格 §7 / §23）。
    /// <para>
    /// 所有 Ribbon callback 都进入这里；异常一律在内部捕获、记录日志并转换为友好消息，
    /// 绝不穿出 COM callback，单个命令失败也不得导致整个插件失效。
    /// </para>
    /// </summary>
    public sealed class CommandDispatcher
    {
        private readonly Dictionary<string, CommandDefinition> _commands =
            new Dictionary<string, CommandDefinition>(StringComparer.OrdinalIgnoreCase);

        private readonly IAccuXContext _context;
        private readonly ILogger _logger;

        public CommandDispatcher(IAccuXContext context, ILogger logger)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _logger = logger ?? NullLogger.Instance;
        }

        public IReadOnlyCollection<CommandDefinition> Commands
        {
            get { return _commands.Values; }
        }

        public void Register(CommandDefinition definition)
        {
            if (definition == null)
            {
                throw new ArgumentNullException(nameof(definition));
            }

            _commands[definition.Id] = definition;
        }

        public void RegisterRange(IEnumerable<CommandDefinition> definitions)
        {
            if (definitions == null)
            {
                return;
            }

            foreach (var definition in definitions)
            {
                Register(definition);
            }
        }

        public bool Contains(string commandId)
        {
            return !string.IsNullOrEmpty(commandId) && _commands.ContainsKey(commandId);
        }

        /// <summary>
        /// 执行命令。任何异常都不会向外抛出。
        /// </summary>
        public CommandResult Execute(string commandId)
        {
            if (!_commands.TryGetValue(commandId ?? string.Empty, out var definition))
            {
                _logger.Warn("未注册的命令: " + commandId);
                return CommandResult.Failed("未找到命令：" + commandId);
            }

            try
            {
                var executionContext = new CommandExecutionContext(definition, _context);
                var result = definition.Handler(executionContext);
                return result ?? CommandResult.Ok();
            }
            catch (Exception ex)
            {
                _logger.Error("命令执行失败: " + commandId, ex);
                return CommandResult.Failed(FriendlyMessage(ex), ex);
            }
        }

        private static string FriendlyMessage(Exception exception)
        {
            if (exception is Operations.HostOperationException)
            {
                return exception.Message;
            }

            return "操作失败：" + exception.Message;
        }
    }
}
