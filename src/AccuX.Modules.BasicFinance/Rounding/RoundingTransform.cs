using AccuX.Core.Cells;
using AccuX.Core.Operations;
using AccuX.Modules.BasicFinance.Common;

namespace AccuX.Modules.BasicFinance.Rounding
{
    /// <summary>
    /// 一键舍入的业务转换（规格 §12 / §15）。
    /// 数值进数值出、公式进公式出；文本/日期/布尔/空白/错误/复杂公式跳过。
    /// </summary>
    public sealed class RoundingTransform : IRangeTransform
    {
        public const string ParameterDigits = "digits";

        private readonly int _digits;

        public RoundingTransform(int digits)
        {
            _digits = digits;
        }

        public string Id
        {
            get { return "rounding"; }
        }

        public bool MayWriteFormula
        {
            get { return true; }
        }

        public CellTransformOutcome Transform(CellData cell, OperationContext context)
        {
            if (cell == null)
            {
                return null;
            }

            if (!cell.Writable)
            {
                return CellTransformOutcome.Skip(cell, SkipReason.NotWritable);
            }

            // 只有“公式计算结果为数值”时才允许包裹 ROUND。
            // FormulaText / FormulaDate / FormulaBoolean / FormulaError 虽然也带有
            // FormulaInfo，但按业务规则必须跳过，不能把字符串、日期或错误值
            // 当作数值公式改写。
            if (cell.CellType == CellValueType.FormulaNumber)
            {
                var formula = cell.Formula;
                if (formula == null || !formula.CanTransform)
                {
                    return CellTransformOutcome.Skip(cell, SkipReason.FormulaNotTransformable);
                }

                var transformed = FormulaTransformService.WrapFunction(
                    formula.Expression,
                    "ROUND",
                    _digits.ToString(System.Globalization.CultureInfo.InvariantCulture));

                return CellTransformOutcome.WriteFormula(cell, transformed);
            }

            // 普通数值：舍入后写回数值。
            if (cell.CellType == CellValueType.ConstantNumber)
            {
                if (!cell.TryGetDecimal(out var value))
                {
                    return CellTransformOutcome.Skip(cell, SkipReason.UnsupportedType);
                }

                var rounded = RoundingService.Round(value, _digits);
                return CellTransformOutcome.WriteValue(cell, rounded);
            }

            return CellTransformOutcome.Skip(cell, MapSkipReason(cell.CellType));
        }

        internal static SkipReason MapSkipReason(CellValueType type)
        {
            switch (type)
            {
                case CellValueType.Blank:
                    return SkipReason.Blank;
                case CellValueType.Text:
                case CellValueType.FormulaText:
                    return SkipReason.Text;
                case CellValueType.Date:
                case CellValueType.FormulaDate:
                    return SkipReason.Date;
                case CellValueType.Boolean:
                case CellValueType.FormulaBoolean:
                    return SkipReason.Boolean;
                case CellValueType.Error:
                case CellValueType.FormulaError:
                    return SkipReason.Error;
                default:
                    return SkipReason.UnsupportedType;
            }
        }
    }
}
