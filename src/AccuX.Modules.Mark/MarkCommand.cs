using System;
using AccuX.Core.Commands;
using AccuX.Core.Logging;
using AccuX.Core.Operations;

namespace AccuX.Modules.Mark
{
    /// <summary>
    /// 对当前选区可见单元格设置指定底色的命令。
    /// </summary>
    public sealed class MarkCommand
    {
        public MarkCommand(string commandId, string displayName, string hexColor)
        {
            if (string.IsNullOrWhiteSpace(commandId))
            {
                throw new ArgumentException("命令 ID 不能为空。", nameof(commandId));
            }

            if (string.IsNullOrWhiteSpace(displayName))
            {
                throw new ArgumentException("显示名称不能为空。", nameof(displayName));
            }

            if (string.IsNullOrWhiteSpace(hexColor))
            {
                throw new ArgumentException("颜色不能为空。", nameof(hexColor));
            }

            CommandId = commandId;
            DisplayName = displayName;
            HexColor = hexColor;
        }

        public string CommandId { get; }

        public string DisplayName { get; }

        public string HexColor { get; }

        public CommandDefinition CreateDefinition(string moduleId)
        {
            return new CommandDefinition(
                CommandId,
                DisplayName,
                moduleId,
                Execute,
                "将当前选区中的可见单元格底色设置为 " + HexColor + "。",
                CommandId);
        }

        private CommandResult Execute(CommandExecutionContext execution)
        {
            var context = execution.Context;
            var host = context.CellMarkHost;
            if (host == null)
            {
                return CommandResult.Failed("当前宿主不支持标记功能。");
            }

            RangeTarget target;
            try
            {
                // 一次命令只捕获一次选区，Host 随后始终依据这个纯 CLR 目标写回。
                target = context.Pipeline.CaptureTarget();
            }
            catch (HostOperationException ex)
            {
                LogOperation(context, execution.Definition.ModuleId, CommandId, null, "capture-failed", ex);
                return CommandResult.Failed(ex.Message, ex);
            }

            try
            {
                var markedCount = host.ApplyVisibleBackgroundColor(target, HexColor);
                LogOperation(context, execution.Definition.ModuleId, CommandId, target, "marked=" + markedCount, null);

                if (markedCount == 0)
                {
                    return CommandResult.Ok("当前选区没有可见单元格，未修改底色。");
                }

                var result = CommandResult.Ok();
                result.ShowMessage = false;
                return result;
            }
            catch (HostOperationException ex)
            {
                LogOperation(context, execution.Definition.ModuleId, CommandId, target, "failed", ex);
                return CommandResult.Failed(ex.Message, ex);
            }
            catch (Exception ex)
            {
                LogOperation(context, execution.Definition.ModuleId, CommandId, target, "failed", ex);
                return CommandResult.Failed("标记选区底色失败：" + ex.Message, ex);
            }
        }

        private static void LogOperation(
            AccuX.Core.Modules.IAccuXContext context,
            string moduleId,
            string commandId,
            RangeTarget target,
            string result,
            Exception exception)
        {
            context.Logger.LogOperation(new OperationLogEntry
            {
                AccuXVersion = context.AccuXVersion,
                HostKind = context.Host?.HostKind.ToString(),
                HostVersion = context.Host?.HostVersion,
                ModuleId = moduleId,
                CommandId = commandId,
                WorksheetName = target?.WorksheetName,
                RangeAddress = target?.Address,
                CellCount = target?.CellCount ?? 0,
                Result = result,
                Exception = exception?.GetType().Name
            });
        }
    }
}
