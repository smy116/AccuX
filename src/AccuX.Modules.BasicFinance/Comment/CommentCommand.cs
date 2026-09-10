using System;
using AccuX.Core.Commands;
using AccuX.Core.Operations;
using AccuX.Modules.BasicFinance.Common;

namespace AccuX.Modules.BasicFinance.Comment
{
    /// <summary>批注助手命令：固定选区目标后读取、编辑、保存或删除传统批注。</summary>
    public sealed class CommentCommand
    {
        public const string CommandId = "accux.basic.comment";

        private readonly ICommentPrompt _prompt;
        private readonly CommentCodec _codec;

        public CommentCommand(ICommentPrompt prompt, CommentCodec codec = null)
        {
            _prompt = prompt ?? throw new ArgumentNullException(nameof(prompt));
            _codec = codec ?? new CommentCodec();
        }

        public CommandDefinition CreateDefinition(string moduleId)
        {
            return new CommandDefinition(
                CommandId,
                "批注助手",
                moduleId,
                Execute,
                "查看、保存或删除当前选区左上角单元格的批注。",
                CommandId);
        }

        private CommandResult Execute(CommandExecutionContext execution)
        {
            var host = execution.Context.CellCommentHost;
            if (host == null)
            {
                return CommandResult.Failed("当前宿主不支持批注助手。");
            }

            CellCommentTarget target;
            try
            {
                target = host.CaptureCommentTarget();
                if (target == null)
                {
                    return CommandResult.Failed("无法确定批注目标，请重新选择一个单元格。");
                }
            }
            catch (HostOperationException ex)
            {
                return CommandResult.Failed(ex.Message, ex);
            }

            CommentContent existing;
            try
            {
                existing = _codec.Decode(host.ReadComment(target));
            }
            catch (HostOperationException ex)
            {
                return CommandResult.Failed(ex.Message, ex);
            }
            catch (Exception ex)
            {
                return CommandResult.Failed("读取单元格批注失败：" + ex.Message, ex);
            }

            if (_prompt is ICommentPromptSession session)
            {
                CommandResult submitted = null;
                var dialogResult = session.AskComment(
                    existing,
                    response =>
                    {
                        submitted = Submit(host, target, existing, response);
                        return submitted.Success
                            ? CommentPromptSubmissionResult.Accepted()
                            : CommentPromptSubmissionResult.Rejected(submitted.Message);
                    });

                if (dialogResult == null || dialogResult.Action == CommentDialogAction.Cancel)
                {
                    return CommandResult.Cancelled();
                }

                return submitted ?? CommandResult.Failed("批注操作未完成。");
            }

            var simpleDialogResult = _prompt.AskComment(existing);
            if (simpleDialogResult == null || simpleDialogResult.Action == CommentDialogAction.Cancel)
            {
                return CommandResult.Cancelled();
            }

            return Submit(host, target, existing, simpleDialogResult);
        }

        private CommandResult Submit(
            ICellCommentHost host,
            CellCommentTarget target,
            CommentContent existing,
            CommentDialogResult dialogResult)
        {
            if (dialogResult == null || dialogResult.Action == CommentDialogAction.Cancel)
            {
                return CommandResult.Cancelled();
            }

            if (dialogResult.Action == CommentDialogAction.Delete)
            {
                try
                {
                    host.DeleteComment(target);
                    return CommandResult.Ok("批注已删除。");
                }
                catch (HostOperationException ex)
                {
                    return CommandResult.Failed(ex.Message, ex);
                }
                catch (Exception ex)
                {
                    return CommandResult.Failed("删除单元格批注失败：" + ex.Message, ex);
                }
            }

            if (dialogResult.Action != CommentDialogAction.Save)
            {
                return CommandResult.Cancelled();
            }

            if (existing.IsUnreadable)
            {
                return CommandResult.Failed(existing.ErrorMessage);
            }

            var text = dialogResult.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text))
            {
                return CommandResult.Failed("批注内容不能为空；如需移除批注，请使用“删除”。");
            }

            var length = CommentTextValidator.CountTextElements(text);
            if (length > CommentTextValidator.MaxLength)
            {
                return CommandResult.Failed(
                    "批注内容超过" + CommentTextValidator.MaxLength + "字限制，请缩短后再保存。");
            }

            string encoded;
            try
            {
                encoded = _codec.Encode(text, dialogResult.Kind);
            }
            catch (Exception ex)
            {
                return CommandResult.Failed("加密批注失败：" + ex.Message, ex);
            }

            try
            {
                host.SaveComment(target, encoded);
                // 保存后编辑窗口即关闭，不再弹出“批注已保存”提示。
                var saved = CommandResult.Ok();
                saved.ShowMessage = false;
                return saved;
            }
            catch (HostOperationException ex)
            {
                return CommandResult.Failed(ex.Message, ex);
            }
            catch (Exception ex)
            {
                return CommandResult.Failed("保存单元格批注失败：" + ex.Message, ex);
            }
        }
    }
}
