using System;
using System.Collections.Generic;

namespace AccuX.Core.Operations
{
    /// <summary>
    /// Host 按 RangeTarget 批量读取并归一化后的完整结果。
    /// </summary>
    public sealed class RangeReadResult
    {
        public RangeReadResult(RangeTarget target, IReadOnlyList<CellData> cells)
        {
            Target = target ?? throw new ArgumentNullException(nameof(target));
            Cells = cells ?? throw new ArgumentNullException(nameof(cells));
            ReadVisibilitySummary();
        }

        public RangeTarget Target { get; }

        /// <summary>按行优先顺序排列的单元格数据，长度等于 Target.CellCount。</summary>
        public IReadOnlyList<CellData> Cells { get; }

        /// <summary>选区当前是否包含被隐藏的行。</summary>
        public bool ContainsHiddenRows { get; private set; }

        /// <summary>选区当前是否包含被隐藏的列。</summary>
        public bool ContainsHiddenColumns { get; private set; }

        /// <summary>选区中同时受隐藏行或隐藏列影响的单元格数量。</summary>
        public int HiddenCellCount { get; private set; }

        private void ReadVisibilitySummary()
        {
            for (var i = 0; i < Cells.Count; i++)
            {
                var cell = Cells[i];
                if (cell == null)
                {
                    continue;
                }

                if (cell.IsHiddenRow)
                {
                    ContainsHiddenRows = true;
                }

                if (cell.IsHiddenColumn)
                {
                    ContainsHiddenColumns = true;
                }

                if (cell.IsHidden)
                {
                    HiddenCellCount++;
                }
            }
        }
    }

    /// <summary>
    /// 单个单元格的待写入指令。
    /// 数值进数值出、公式进公式出（规格 §11）：两者互斥。
    /// </summary>
    public sealed class CellWrite
    {
        private CellWrite(int row, int column)
        {
            Row = row;
            Column = column;
        }

        public int Row { get; }

        public int Column { get; }

        /// <summary>待写入的数值/文本/布尔/日期值；与 Formula 互斥。</summary>
        public object Value { get; private set; }

        /// <summary>待写入的规范化公式表达式（以 '=' 开头）；与 Value 互斥。</summary>
        public string Formula { get; private set; }

        /// <summary>可选的目标 NumberFormat；null 表示不修改。</summary>
        public string NumberFormat { get; private set; }

        public bool IsFormulaWrite
        {
            get { return Formula != null; }
        }

        public static CellWrite ValueWrite(int row, int column, object value, string numberFormat = null)
        {
            return new CellWrite(row, column)
            {
                Value = value,
                NumberFormat = numberFormat
            };
        }

        public static CellWrite FormulaWrite(int row, int column, string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
            {
                throw new ArgumentException("公式表达式不能为空。", nameof(formula));
            }

            return new CellWrite(row, column)
            {
                Formula = formula
            };
        }
    }

    /// <summary>
    /// 完整待写结果（规格 §10 / §16）：在开始写回前于内存中生成。
    /// </summary>
    public sealed class RangeWritePlan
    {
        private readonly List<CellWrite> _writes = new List<CellWrite>();

        public RangeWritePlan(RangeTarget target)
        {
            Target = target ?? throw new ArgumentNullException(nameof(target));
        }

        public RangeTarget Target { get; }

        public IReadOnlyList<CellWrite> Writes
        {
            get { return _writes; }
        }

        public bool IsEmpty
        {
            get { return _writes.Count == 0; }
        }

        /// <summary>需要写回 NumberFormat 的单元格。</summary>
        public void Add(CellWrite write)
        {
            if (write == null)
            {
                throw new ArgumentNullException(nameof(write));
            }

            _writes.Add(write);
        }

        public void AddValue(int row, int column, object value, string numberFormat = null)
        {
            Add(CellWrite.ValueWrite(row, column, value, numberFormat));
        }

        public void AddFormula(int row, int column, string formula)
        {
            Add(CellWrite.FormulaWrite(row, column, formula));
        }
    }
}
