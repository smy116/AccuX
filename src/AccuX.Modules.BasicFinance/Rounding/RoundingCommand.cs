using System;
using System.Collections.Generic;
using AccuX.Core.Commands;
using AccuX.Core.Operations;
using AccuX.Modules.BasicFinance.Common;

namespace AccuX.Modules.BasicFinance.Rounding
{
    /// <summary>
    /// 一键舍入命令（规格 §12 / §20）。
    /// 每次点击弹出小数位窗口（默认预填 2 位），确认后执行。
    /// </summary>
    public sealed class RoundingCommand
    {
        public const string CommandId = "accux.basic.round";

        private readonly IUserPrompt _prompt;
        private readonly HostStateOptions _stateOptions;

        public RoundingCommand(IUserPrompt prompt, HostStateOptions stateOptions = null)
        {
            _prompt = prompt;
            _stateOptions = stateOptions ?? new HostStateOptions();
        }

        public CommandDefinition CreateDefinition(string moduleId)
        {
            return new CommandDefinition(
                CommandId,
                "一键舍入",
                moduleId,
                Execute,
                "对选区金额统一四舍五入。",
                CommandId);
        }

        private CommandResult Execute(CommandExecutionContext execution)
        {
            var context = execution.Context;

            // 1. 唯一一次捕获 RangeTarget（规格 §10 / §19）。
            RangeTarget rangeTarget;
            try
            {
                rangeTarget = context.Pipeline.CaptureTarget();
            }
            catch (HostOperationException ex)
            {
                return CommandResult.Failed(ex.Message, ex);
            }

            // 2. 弹出参数窗口。
            var options = _prompt.AskRoundingOptions(rangeTarget);
            if (options == null)
            {
                return CommandResult.Cancelled();
            }

            // 3. 大选区确认由 Host CaptureTarget 负责硬限制，这里给出确认提示。
            if (!_prompt.ConfirmLargeSelection(rangeTarget))
            {
                return CommandResult.Cancelled();
            }

            var operationContext = new OperationContext(CommandId, execution.Definition.ModuleId, rangeTarget);
            var transform = new RoundingTransform(options.Digits);
            var result = context.Pipeline.Execute(transform, operationContext, _stateOptions);

            return ToCommandResult(result);
        }

        internal static CommandResult ToCommandResult(OperationResult result)
        {
            if (result == null)
            {
                return CommandResult.Failed("操作未返回结果。");
            }

            if (!result.Success)
            {
                return CommandResult.Failed(result.Message);
            }

            return CommandResult.Ok(result.Message);
        }
    }
}
