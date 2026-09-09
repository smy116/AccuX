using AccuX.Core.Commands;
using AccuX.Core.Operations;
using AccuX.Modules.BasicFinance.AmountConversion;
using AccuX.Modules.BasicFinance.Common;

namespace AccuX.Modules.BasicFinance.AmountConversion
{
    /// <summary>
    /// 金额折合命令（规格 §13 / §20）。
    /// 每次点击弹出折合参数窗口，确认后执行。
    /// </summary>
    public sealed class AmountConversionCommand
    {
        public const string CommandId = "accux.basic.convert";

        private readonly IUserPrompt _prompt;
        private readonly HostStateOptions _stateOptions;

        public AmountConversionCommand(IUserPrompt prompt, HostStateOptions stateOptions = null)
        {
            _prompt = prompt;
            _stateOptions = stateOptions ?? new HostStateOptions();
        }

        public CommandDefinition CreateDefinition(string moduleId)
        {
            return new CommandDefinition(
                CommandId,
                "金额折合",
                moduleId,
                Execute,
                "按指定折合率对选区金额进行折合。",
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

            var options = _prompt.AskAmountConversionOptions(rangeTarget);
            if (options == null)
            {
                return CommandResult.Cancelled();
            }

            if (!AmountConversionOptions.TryValidate(options, out var error))
            {
                _prompt.ShowError(error);
                return CommandResult.Failed(error);
            }

            if (!_prompt.ConfirmLargeSelection(rangeTarget))
            {
                return CommandResult.Cancelled();
            }

            var operationContext = new OperationContext(CommandId, execution.Definition.ModuleId, rangeTarget);
            var transform = new AmountConversionTransform(options);
            var result = context.Pipeline.Execute(transform, operationContext, _stateOptions);

            return Rounding.RoundingCommand.ToCommandResult(result);
        }
    }
}
