using System;

namespace AccuX.Core.Logging
{
    /// <summary>
    /// 不写入任何内容的日志实现，用于单元测试和默认装配。
    /// </summary>
    public sealed class NullLogger : ILogger
    {
        public static readonly NullLogger Instance = new NullLogger();

        public void Debug(string message)
        {
        }

        public void Info(string message)
        {
        }

        public void Warn(string message)
        {
        }

        public void Error(string message, Exception exception = null)
        {
        }

        public void LogOperation(OperationLogEntry entry)
        {
        }
    }
}
