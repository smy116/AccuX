using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace AccuX.Core.Logging
{
    /// <summary>
    /// 按天写入本地文本日志。所有写入都包裹在 try/catch 中，日志失败绝不向上抛出。
    /// </summary>
    public sealed class FileLogger : ILogger
    {
        private readonly object _gate = new object();
        private readonly string _directory;
        private readonly LogLevel _minimumLevel;

        public FileLogger(string directory, LogLevel minimumLevel = LogLevel.Info)
        {
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new ArgumentException("日志目录不能为空。", nameof(directory));
            }

            _directory = directory;
            _minimumLevel = minimumLevel;
        }

        public static string DefaultDirectory
        {
            get
            {
                var root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                return Path.Combine(root, "AccuX", "logs");
            }
        }

        public void Debug(string message)
        {
            Write(LogLevel.Debug, message, null);
        }

        public void Info(string message)
        {
            Write(LogLevel.Info, message, null);
        }

        public void Warn(string message)
        {
            Write(LogLevel.Warn, message, null);
        }

        public void Error(string message, Exception exception = null)
        {
            Write(LogLevel.Error, message, exception);
        }

        public void LogOperation(OperationLogEntry entry)
        {
            if (entry == null)
            {
                return;
            }

            var builder = new StringBuilder();
            builder.Append("command=").Append(entry.CommandId);
            builder.Append(" module=").Append(entry.ModuleId);
            builder.Append(" host=").Append(entry.HostKind).Append(' ').Append(entry.HostVersion);
            builder.Append(" sheet=").Append(entry.WorksheetName);
            builder.Append(" range=").Append(entry.RangeAddress);
            builder.Append(" cells=").Append(entry.CellCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" durationMs=").Append(((long)entry.Duration.TotalMilliseconds).ToString(CultureInfo.InvariantCulture));
            builder.Append(" result=").Append(entry.Result);
            builder.Append(" accux=").Append(entry.AccuXVersion);

            if (!string.IsNullOrEmpty(entry.Exception))
            {
                builder.Append(" exception=").Append(entry.Exception);
            }

            Write(LogLevel.Info, builder.ToString(), null);
        }

        private void Write(LogLevel level, string message, Exception exception)
        {
            if (level < _minimumLevel)
            {
                return;
            }

            try
            {
                var line = new StringBuilder();
                line.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture));
                line.Append(' ').Append(level.ToString().ToUpperInvariant().PadRight(5));
                line.Append(' ').Append(message ?? string.Empty);

                if (exception != null)
                {
                    line.Append(" | ").Append(exception.GetType().Name).Append(": ").Append(exception.Message);
                }

                lock (_gate)
                {
                    Directory.CreateDirectory(_directory);
                    var path = Path.Combine(_directory, "accux-" + DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + ".log");
                    File.AppendAllText(path, line.ToString() + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch
            {
                // 日志失败不得影响业务执行。
            }
        }
    }
}
