using System;

namespace AccuX.Core.Operations
{
    /// <summary>
    /// 一次 Command 的唯一写入目标描述（规格 §5.2 / §19）。
    /// <para>
    /// 该类型只包含 CLR 身份信息与地址信息，<b>严禁</b>携带 Workbook / Worksheet / Range 等 COM 对象。
    /// WorkbookKey / WorksheetKey 由 Host 生成并解释，Core 与业务 Module 不解释其含义。
    /// </para>
    /// </summary>
    public sealed class RangeTarget
    {
        public RangeTarget(
            string workbookKey,
            string worksheetKey,
            string worksheetName,
            string address,
            int rowCount,
            int columnCount,
            long cellCount,
            bool isMultiArea,
            bool containsMergedCells,
            string identityToken = null)
        {
            WorkbookKey = workbookKey ?? throw new ArgumentNullException(nameof(workbookKey));
            WorksheetKey = worksheetKey ?? throw new ArgumentNullException(nameof(worksheetKey));
            WorksheetName = worksheetName ?? string.Empty;
            Address = address ?? throw new ArgumentNullException(nameof(address));
            RowCount = rowCount;
            ColumnCount = columnCount;
            CellCount = cellCount;
            IsMultiArea = isMultiArea;
            ContainsMergedCells = containsMergedCells;
            IdentityToken = identityToken ?? string.Empty;
        }

        /// <summary>由 Host 生成和解释的稳定 Workbook 身份信息。</summary>
        public string WorkbookKey { get; }

        /// <summary>由 Host 生成和解释的稳定 Worksheet 身份信息。</summary>
        public string WorksheetKey { get; }

        /// <summary>工作表名称，仅用于日志与提示，不作为身份判断依据。</summary>
        public string WorksheetName { get; }

        /// <summary>选区地址，例如 A1:C10。</summary>
        public string Address { get; }

        public int RowCount { get; }

        public int ColumnCount { get; }

        public long CellCount { get; }

        /// <summary>是否为多区域 Selection。</summary>
        public bool IsMultiArea { get; }

        /// <summary>是否包含合并单元格。</summary>
        public bool ContainsMergedCells { get; }

        /// <summary>
        /// Host 会话内的工作簿身份令牌。空值表示使用传统 WorkbookKey 解析。
        /// </summary>
        public string IdentityToken { get; }

        public override string ToString()
        {
            return WorksheetName + "!" + Address + " (" + CellCount + " cells)";
        }
    }
}
