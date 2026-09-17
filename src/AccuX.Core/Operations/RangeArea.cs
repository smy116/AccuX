using System;

namespace AccuX.Core.Operations
{
    /// <summary>选区中的一个连续矩形，不包含 COM 引用。</summary>
    public sealed class RangeArea
    {
        public RangeArea(string address, int rowCount, int columnCount)
        {
            if (string.IsNullOrWhiteSpace(address)) throw new ArgumentException("区域地址不能为空。", nameof(address));
            if (rowCount <= 0) throw new ArgumentOutOfRangeException(nameof(rowCount));
            if (columnCount <= 0) throw new ArgumentOutOfRangeException(nameof(columnCount));
            Address = address;
            RowCount = rowCount;
            ColumnCount = columnCount;
        }

        public string Address { get; }
        public int RowCount { get; }
        public int ColumnCount { get; }
        public long CellCount => (long)RowCount * ColumnCount;
    }
}
