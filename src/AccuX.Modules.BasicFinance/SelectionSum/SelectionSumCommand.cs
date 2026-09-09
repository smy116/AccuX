using System;
using System.Diagnostics;
using AccuX.Core.Commands;
using AccuX.Core.Logging;
using AccuX.Core.Operations;
using AccuX.Modules.BasicFinance.Common;

namespace AccuX.Modules.BasicFinance.SelectionSum
{
    /// <summary>
    /// 选区求和命令：只读当前选区，将可见数字合计复制到剪切板。
    /// </summary>
    public sealed class SelectionSumCommand
    {
        public const string CommandId = "accux.basic.sum";

        private readonly IUserPrompt _prompt;

        public SelectionSumCommand(IUserPrompt prompt)
        {
            _prompt = prompt ?? throw new ArgumentNullException(nameof(prompt));
        }

        public CommandDefinition CreateDefinition(string moduleId)
        {
            return new CommandDefinition(
                CommandId,
                "选区求和",
                moduleId,
                Execute,
                "计算当前选区内可见数字合计并复制到剪切板。",
                CommandId);
        }

        private CommandResult Execute(CommandExecutionContext execution)
        {
            var context = execution.Context;

            RangeTarget rangeTarget;
            try
            {
                rangeTarget = context.Pipeline.CaptureTarget();
            }
            catch (HostOperationException ex)
            {
                return CommandResult.Failed(ex.Message, ex);
            }

            if (!_prompt.ConfirmLargeSelection(rangeTarget))
            {
                return CommandResult.Cancelled();
            }

            var stopwatch = Stopwatch.StartNew();
            try
            {
                var readResult = context.Pipeline.Read(rangeTarget);
                var sum = SelectionSumService.Calculate(readResult.Cells);

                if (!_prompt.TryCopyToClipboard(sum.FormattedTotal))
                {
                    stopwatch.Stop();
                    LogOperation(context, execution.Definition.ModuleId, rangeTarget, stopwatch.Elapsed, "copy-failed", null);
                    return CommandResult.Failed("合计已计算，但复制至剪切板失败，请重试。");
                }

                stopwatch.Stop();
                LogOperation(context, execution.Definition.ModuleId, rangeTarget, stopwatch.Elapsed, "copied", null);

                return CommandResult.Ok("选定区域合计" + sum.FormattedTotal + "，已复制至剪切板。");
            }
            catch (HostOperationException ex)
            {
                stopwatch.Stop();
                LogOperation(context, execution.Definition.ModuleId, rangeTarget, stopwatch.Elapsed, "failed", ex);
                return CommandResult.Failed(ex.Message, ex);
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                LogOperation(context, execution.Definition.ModuleId, rangeTarget, stopwatch.Elapsed, "failed", ex);
                return CommandResult.Failed("操作失败：" + ex.Message, ex);
            }
        }

        private static void LogOperation(
            AccuX.Core.Modules.IAccuXContext context,
            string moduleId,
            RangeTarget target,
            TimeSpan duration,
            string result,
            Exception exception)
        {
            context.Logger.LogOperation(new OperationLogEntry
            {
                AccuXVersion = context.AccuXVersion,
                HostKind = context.Host?.HostKind.ToString(),
                HostVersion = context.Host?.HostVersion,
                ModuleId = moduleId,
                CommandId = CommandId,
                WorksheetName = target.WorksheetName,
                RangeAddress = target.Address,
                CellCount = target.CellCount,
                Duration = duration,
                Result = result,
                Exception = exception?.GetType().Name
            });
        }
    }
}
