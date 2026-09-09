using System;
using System.Globalization;

namespace AccuX.Core.Cells
{
    /// <summary>
    /// 单元格分类的宿主输入。
    /// Host 负责填充已验证的归一化信息，分类逻辑集中在 <see cref="CellValueClassifier"/>。
    /// </summary>
    public sealed class CellClassificationInput
    {
        /// <summary>单元格是否为空白（宿主判定）。</summary>
        public bool IsBlank { get; set; }

        /// <summary>单元格是否包含公式（宿主判定）。</summary>
        public bool HasFormula { get; set; }

        /// <summary>
        /// 归一化后的计算结果值：decimal / string / bool / DateTime / null / 错误文本。
        /// </summary>
        public object Value { get; set; }

        /// <summary>宿主是否将结果识别为日期/时间。</summary>
        public bool IsDate { get; set; }

        /// <summary>宿主是否将结果识别为错误值。</summary>
        public bool IsError { get; set; }

        /// <summary>数值是否由文本格式伪装（例如前置单引号）。</summary>
        public bool IsText { get; set; }
    }

    /// <summary>
    /// 唯一的单元格分类入口（规格 §9）。
    /// 纯 CLR 实现，可脱离 Excel/WPS 单元测试。
    /// </summary>
    public static class CellValueClassifier
    {
        /// <summary>
        /// 根据宿主归一化输入输出统一分类。
        /// </summary>
        public static CellValueType Classify(CellClassificationInput input)
        {
            if (input == null)
            {
                return CellValueType.Blank;
            }

            var isFormula = input.HasFormula;

            if (input.IsError)
            {
                return isFormula ? CellValueType.FormulaError : CellValueType.Error;
            }

            if (input.IsBlank && input.Value == null)
            {
                return CellValueType.Blank;
            }

            if (input.Value == null)
            {
                return CellValueType.Blank;
            }

            if (input.IsDate || input.Value is DateTime)
            {
                return isFormula ? CellValueType.FormulaDate : CellValueType.Date;
            }

            if (!input.IsText && input.Value is bool)
            {
                return isFormula ? CellValueType.FormulaBoolean : CellValueType.Boolean;
            }

            if (!input.IsText && IsNumeric(input.Value))
            {
                return isFormula ? CellValueType.FormulaNumber : CellValueType.ConstantNumber;
            }

            if (input.Value is string)
            {
                return isFormula ? CellValueType.FormulaText : CellValueType.Text;
            }

            return CellValueType.Unsupported;
        }

        /// <summary>
        /// 将宿主原始值归一化为 CLR 值（decimal 优先，保证财务精度）。
        /// </summary>
        public static object NormalizeNumeric(object raw)
        {
            switch (raw)
            {
                case null:
                    return null;
                case decimal d:
                    return d;
                case double dbl:
                    return (decimal)dbl;
                case float flt:
                    return (decimal)flt;
                case int i:
                    return (decimal)i;
                case long l:
                    return (decimal)l;
                case short s:
                    return (decimal)s;
                case byte b:
                    return (decimal)b;
                case string str:
                    return decimal.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed) ? (object)parsed : str;
                default:
                    return raw;
            }
        }

        private static bool IsNumeric(object value)
        {
            switch (value)
            {
                case decimal _:
                case double _:
                case float _:
                case int _:
                case long _:
                case short _:
                case byte _:
                    return true;
                default:
                    return false;
            }
        }
    }
}
