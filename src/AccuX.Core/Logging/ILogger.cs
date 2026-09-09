using System;

namespace AccuX.Core.Logging
{
    /// <summary>
    /// 日志级别。
    /// </summary>
    public enum LogLevel
    {
        Debug = 0,
        Info = 1,
        Warn = 2,
        Error = 3
    }

    /// <summary>
    /// AccuX 日志抽象。实现必须保证：写入失败不得抛出异常影响业务。
    /// </summary>
    public interface ILogger
    {
        void Debug(string message);

        void Info(string message);

        void Warn(string message);

        void Error(string message, Exception exception = null);

        /// <summary>
        /// 记录一次业务操作的结构化日志（见规格 §22）。
        /// 该入口不得记录真实金额、公式内容或其他敏感数据。
        /// </summary>
        void LogOperation(OperationLogEntry entry);
    }
}
