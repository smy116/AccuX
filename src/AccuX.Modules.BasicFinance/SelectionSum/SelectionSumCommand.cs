using System;
using System.Diagnostics;
using AccuX.Core.Commands;
using AccuX.Core.Logging;
using AccuX.Core.Operations;
using AccuX.Modules.BasicFinance.Common;

namespace AccuX.Modules.BasicFinance.SelectionSum
{
    /// <summary>
    /// 选区求和命令：只读当前选区，显示金额的多格式复制窗口。
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
                "计算当前选区内可见数字合计，提供金额、万元金额和大写金额复制。",
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
                var dialogResult = _prompt.ShowSelectionSumDialog(sum);

                if (dialogResult == null || dialogResult.WasCancelled)
                {
                    stopwatch.Stop();
                    LogOperation(context, execution.Definition.ModuleId, rangeTarget, stopwatch.Elapsed, "cancelled", null);
                    return CommandResult.Cancelled();
                }

                stopwatch.Stop();
                var result = dialogResult.CopySucceeded
                    ? "copied-" + GetCopyKindLogName(dialogResult.CopyKind)
                    : "copy-failed-" + GetCopyKindLogName(dialogResult.CopyKind);
                LogOperation(context, execution.Definition.ModuleId, rangeTarget, stopwatch.Elapsed, result, null);

                if (!dialogResult.CopySucceeded)
                {
                    var failure = CommandResult.Failed("复制至剪切板失败，请重试。");
                    failure.ShowMessage = false;
                    return failure;
                }

                var success = CommandResult.Ok();
                success.ShowMessage = false;
                return success;
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

        private static string GetCopyKindLogName(SelectionSumCopyKind copyKind)
        {
            switch (copyKind)
            {
                case SelectionSumCopyKind.Amount:
                    return "amount";
                case SelectionSumCopyKind.WanAmount:
                    return "wan-amount";
                case SelectionSumCopyKind.ChineseAmount:
                    return "chinese-amount";
                default:
                    return "unknown";
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
