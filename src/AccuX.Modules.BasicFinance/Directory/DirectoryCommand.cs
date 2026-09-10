using System;
using System.Globalization;
using AccuX.Core.Commands;
using AccuX.Core.Operations;
using AccuX.Modules.BasicFinance.Common;

namespace AccuX.Modules.BasicFinance.Directory
{
    /// <summary>
    /// 生成当前工作簿目录的命令：决定目录参数与替换策略，渲染交由宿主完成。
    /// </summary>
    public sealed class DirectoryCommand
    {
        public const string CommandId = "accux.basic.directory";

        private readonly IUserPrompt _prompt;

        public DirectoryCommand(IUserPrompt prompt)
        {
            _prompt = prompt ?? throw new ArgumentNullException(nameof(prompt));
        }

        public CommandDefinition CreateDefinition(string moduleId)
        {
            return new CommandDefinition(
                CommandId,
                "生成目录",
                moduleId,
                Execute,
                "在当前工作簿最前面生成可见工作表目录。",
                CommandId);
        }

        private CommandResult Execute(CommandExecutionContext execution)
        {
            var host = execution.Context.WorkbookDirectoryHost;
            if (host == null)
            {
                return CommandResult.Failed("当前宿主不支持生成目录。");
            }

            try
            {
                var options = DirectoryTemplate.CreateOptions();

                var replaceExisting = false;
                if (host.DirectoryWorksheetExists(options.WorksheetName))
                {
                    if (!_prompt.ConfirmReplaceDirectory(options.WorksheetName))
                    {
                        return CommandResult.Cancelled();
                    }

                    replaceExisting = true;
                }

                var count = host.GenerateDirectory(options, replaceExisting);
                var message = string.Format(
                    CultureInfo.InvariantCulture,
                    "生成完成！共生成{0}个表格的目录。",
                    count);

                return CommandResult.Ok(message);
            }
            catch (HostOperationException ex)
            {
                return CommandResult.Failed(ex.Message, ex);
            }
            catch (Exception ex)
            {
                return CommandResult.Failed("操作失败：" + ex.Message, ex);
            }
        }
    }
}
