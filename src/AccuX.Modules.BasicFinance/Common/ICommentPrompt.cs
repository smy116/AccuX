using System;
using AccuX.Modules.BasicFinance.Comment;

namespace AccuX.Modules.BasicFinance.Common
{
    /// <summary>批注编辑窗口的交互抽象，便于命令脱离 WPF 进行测试。</summary>
    public interface ICommentPrompt
    {
        CommentDialogResult AskComment(CommentContent existing);
    }

    /// <summary>
    /// 支持提交回调的批注窗口交互。提交失败时窗口可以保留输入并继续编辑。
    /// </summary>
    public interface ICommentPromptSession : ICommentPrompt
    {
        CommentDialogResult AskComment(
            CommentContent existing,
            Func<CommentDialogResult, CommentPromptSubmissionResult> submit);
    }

    public sealed class CommentPromptSubmissionResult
    {
        private CommentPromptSubmissionResult(bool success, string message)
        {
            Success = success;
            Message = message ?? string.Empty;
        }

        public bool Success { get; }

        public string Message { get; }

        public static CommentPromptSubmissionResult Accepted()
        {
            return new CommentPromptSubmissionResult(true, string.Empty);
        }

        public static CommentPromptSubmissionResult Rejected(string message)
        {
            return new CommentPromptSubmissionResult(false, message);
        }
    }

    public sealed class CommentDialogResult
    {
        private CommentDialogResult(CommentDialogAction action, string text, CommentKind kind)
        {
            Action = action;
            Text = text ?? string.Empty;
            Kind = kind;
        }

        public CommentDialogAction Action { get; }

        public string Text { get; }

        public CommentKind Kind { get; }

        public static CommentDialogResult Cancelled()
        {
            return new CommentDialogResult(CommentDialogAction.Cancel, string.Empty, CommentKind.Plain);
        }

        public static CommentDialogResult Save(string text, CommentKind kind)
        {
            return new CommentDialogResult(CommentDialogAction.Save, text, kind);
        }

        public static CommentDialogResult Delete()
        {
            return new CommentDialogResult(CommentDialogAction.Delete, string.Empty, CommentKind.Plain);
        }
    }
}
