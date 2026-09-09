using System;

namespace AccuX.Core.Operations
{
    /// <summary>
    /// 一次批注操作锁定的单元格目标。
    /// 该类型只保存 CLR 身份与地址信息，不携带任何宿主 COM 对象。
    /// </summary>
    public sealed class CellCommentTarget
    {
        public CellCommentTarget(
            string workbookKey,
            string worksheetKey,
            string worksheetName,
            string address,
            int row = 0,
            int column = 0)
        {
            WorkbookKey = workbookKey ?? throw new ArgumentNullException(nameof(workbookKey));
            WorksheetKey = worksheetKey ?? throw new ArgumentNullException(nameof(worksheetKey));
            WorksheetName = worksheetName ?? string.Empty;
            Address = address ?? throw new ArgumentNullException(nameof(address));
            Row = row;
            Column = column;
        }

        public string WorkbookKey { get; }

        public string WorksheetKey { get; }

        public string WorksheetName { get; }

        /// <summary>目标单元格地址，例如 $B$3。</summary>
        public string Address { get; }

        /// <summary>目标单元格的 1-based 行号；宿主无法提供时为 0。</summary>
        public int Row { get; }

        /// <summary>目标单元格的 1-based 列号；宿主无法提供时为 0。</summary>
        public int Column { get; }

        public override string ToString()
        {
            return WorksheetName + "!" + Address;
        }
    }
}
