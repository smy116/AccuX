using AccuX.Core.Commands;
using AccuX.Core.Operations;
using AccuX.Modules.BasicFinance.Common;

namespace AccuX.Modules.BasicFinance.ChineseAmount
{
    /// <summary>
    /// 金额大写命令（规格 §14）。
    /// 无参数，直接原地覆盖当前选区。
    /// </summary>
    public sealed class ChineseAmountCommand
    {
        public const string CommandId = "accux.basic.uppercase";

        private readonly IUserPrompt _prompt;
        private readonly HostStateOptions _stateOptions;

        public ChineseAmountCommand(IUserPrompt prompt, HostStateOptions stateOptions = null)
        {
            _prompt = prompt;
            _stateOptions = stateOptions ?? new HostStateOptions();
        }

        public CommandDefinition CreateDefinition(string moduleId)
        {
            return new CommandDefinition(
                CommandId,
                "金额大写",
                moduleId,
                Execute,
                "将选区金额转换为中文大写。",
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

            var operationContext = new OperationContext(CommandId, execution.Definition.ModuleId, rangeTarget);
            var transform = new ChineseAmountTransform();
            var result = context.Pipeline.Execute(transform, operationContext, _stateOptions);

            return Rounding.RoundingCommand.ToCommandResult(result);
        }
    }
}
