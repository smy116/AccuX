using System;

namespace AccuX.Core.Logging
{
    /// <summary>
    /// 一次业务操作的结构化日志条目。
    /// 字段严格遵循规格 §22：只记录可诊断信息，不包含金额、公式或工作表敏感数据。
    /// </summary>
    public sealed class OperationLogEntry
    {
        public string AccuXVersion { get; set; }

        public string HostKind { get; set; }

        public string HostVersion { get; set; }

        public string ModuleId { get; set; }

        public string CommandId { get; set; }

        public string WorksheetName { get; set; }

        public string RangeAddress { get; set; }

        public long CellCount { get; set; }

        public TimeSpan Duration { get; set; }

        public string Result { get; set; }

        public string Exception { get; set; }
    }
}
