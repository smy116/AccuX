using System;
using System.Collections.Generic;
using AccuX.Core.Commands;
using AccuX.Core.Modules;

namespace AccuX.Modules.Mark
{
    /// <summary>
    /// 标记模块：提供四种固定颜色的选区底色标记命令。
    /// </summary>
    public sealed class MarkModule : IAccuXModule
    {
        public const string ModuleId = "AccuX.Modules.Mark";

        private CommandDefinition[] _commands = Array.Empty<CommandDefinition>();

        public string Id
        {
            get { return ModuleId; }
        }

        public void Initialize(IAccuXContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            _commands = new[]
            {
                new MarkCommand("accux.mark.green", "标绿", "#18be6a").CreateDefinition(ModuleId),
                new MarkCommand("accux.mark.red", "标红", "#ed4015").CreateDefinition(ModuleId),
                new MarkCommand("accux.mark.yellow", "标黄", "#fe9900").CreateDefinition(ModuleId),
                new MarkCommand("accux.mark.blue", "标蓝", "#2db7f5").CreateDefinition(ModuleId)
            };

            context.Logger.Info("Mark 模块初始化完成。");
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
