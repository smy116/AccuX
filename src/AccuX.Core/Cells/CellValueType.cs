namespace AccuX.Core.Cells
{
    /// <summary>
    /// 统一的单元格分类结果（规格 §9）。
    /// 四个 AccuX 功能都依据该分类决定处理策略，禁止各自判断空值/公式/数字/日期/布尔。
    /// </summary>
    public enum CellValueType
    {
        /// <summary>空单元格。</summary>
        Blank,

        /// <summary>普通数值常量。</summary>
        ConstantNumber,

        /// <summary>公式，结果为数值。</summary>
        FormulaNumber,

        /// <summary>文本常量。</summary>
        Text,

        /// <summary>公式，结果为文本。</summary>
        FormulaText,

        /// <summary>日期/时间常量（含宿主格式识别的日期序列值）。</summary>
        Date,

        /// <summary>公式，结果为日期/时间。</summary>
        FormulaDate,

        /// <summary>逻辑值常量。</summary>
        Boolean,

        /// <summary>公式，结果为逻辑值。</summary>
        FormulaBoolean,

        /// <summary>错误值常量。</summary>
        Error,

        /// <summary>公式，结果为错误值。</summary>
        FormulaError,

        /// <summary>无法识别或不支持安全处理的类型。</summary>
        Unsupported
    }

    public static class CellValueTypeExtensions
    {
        /// <summary>是否为公式单元格（任意公式子类）。</summary>
        public static bool IsFormula(this CellValueType type)
        {
            switch (type)
            {
                case CellValueType.FormulaNumber:
                case CellValueType.FormulaText:
                case CellValueType.FormulaDate:
                case CellValueType.FormulaBoolean:
                case CellValueType.FormulaError:
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>是否为数值（普通数值或数值公式）。</summary>
        public static bool IsNumber(this CellValueType type)
        {
            return type == CellValueType.ConstantNumber || type == CellValueType.FormulaNumber;
        }
    }
}
