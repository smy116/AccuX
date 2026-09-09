using System;

namespace AccuX.Core.Commands
{
    /// <summary>
    /// 命令执行结果。
    /// </summary>
    public sealed class CommandResult
    {
        public CommandResult(bool success, string message, Exception exception = null)
        {
            Success = success;
            Message = message ?? string.Empty;
            Exception = exception;
        }

        public bool Success { get; }

        public string Message { get; }

        public Exception Exception { get; }

        /// <summary>是否需要在 UI 上向用户展示消息。</summary>
        public bool ShowMessage { get; set; } = true;

        public static CommandResult Ok(string message = null)
        {
            return new CommandResult(true, message);
        }

        public static CommandResult Failed(string message, Exception exception = null)
        {
            return new CommandResult(false, message, exception);
        }

        public static CommandResult Cancelled()
        {
            return new CommandResult(true, null) { ShowMessage = false };
        }
    }
}
