using System;
using System.Globalization;

namespace AccuX.Core.Operations
{
    /// <summary>
    /// Host 批量读取后归一化得到的单个单元格数据。
    /// 只包含 CLR 数据，不包含任何 COM 引用。
    /// </summary>
    public sealed class CellData
    {
        public CellData(int row, int column, int areaIndex = 0)
        {
            AreaIndex = areaIndex;
            Row = row;
            Column = column;
        }

        public int AreaIndex { get; }

        /// <summary>相对于所属区域的行偏移（0 基）。</summary>
        public int Row { get; }

        /// <summary>相对于所属区域的列偏移（0 基）。</summary>
        public int Column { get; }

        /// <summary>统一分类结果。</summary>
        public Cells.CellValueType CellType { get; set; }

        /// <summary>
        /// 归一化后的 CLR 值：
        /// decimal（数值）、string（文本）、bool、DateTime、null（空）、或 Error 的字符串表示。
        /// 公式单元格此处为其计算结果值。
        /// </summary>
        public object Value { get; set; }

        /// <summary>公式信息；非公式单元格为 null。</summary>
        public FormulaInfo Formula { get; set; }

        /// <summary>宿主 NumberFormat 字符串，用于日期/文本识别。</summary>
        public string NumberFormat { get; set; }

        /// <summary>是否属于合并单元格。</summary>
        public bool IsMerged { get; set; }

        /// <summary>所在行当前是否被宿主隐藏。</summary>
        public bool IsHiddenRow { get; set; }

        /// <summary>所在列当前是否被宿主隐藏。</summary>
        public bool IsHiddenColumn { get; set; }

        /// <summary>当前单元格是否因所在行或列被宿主隐藏。</summary>
        public bool IsHidden
        {
            get { return IsHiddenRow || IsHiddenColumn; }
        }

        /// <summary>该单元格当前是否允许写入。</summary>
        public bool Writable { get; set; }

        public bool IsFormula
        {
            get { return Formula != null; }
        }

        /// <summary>
        /// 尝试将当前值转换为 decimal。仅对数值类型有效。
        /// </summary>
        public bool TryGetDecimal(out decimal value)
        {
            value = 0m;

            if (Value == null)
            {
                return false;
            }

            switch (CellType)
            {
                case Cells.CellValueType.ConstantNumber:
                case Cells.CellValueType.FormulaNumber:
                    break;
                default:
                    return false;
            }

            switch (Value)
            {
                case decimal d:
                    value = d;
                    return true;
                case double dbl:
                    value = (decimal)dbl;
                    return true;
                case float flt:
                    value = (decimal)flt;
                    return true;
                case int i:
                    value = i;
                    return true;
                case long l:
                    value = l;
                    return true;
                case short s:
                    value = s;
                    return true;
                case byte b:
                    value = b;
                    return true;
                case string str:
                    return decimal.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out value);
                default:
                    return false;
            }
        }

        public override string ToString()
        {
            return "[" + Row + "," + Column + "] " + CellType + "=" + (Value ?? "<null>");
        }
    }
}
