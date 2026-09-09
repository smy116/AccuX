using AccuX.Core.Cells;
using AccuX.Core.Operations;
using AccuX.Modules.BasicFinance.Common;

namespace AccuX.Modules.BasicFinance.AmountConversion
{
    /// <summary>
    /// 金额折合的业务转换（规格 §13 / §15）。
    /// 数值进数值出、公式进公式出。
    /// </summary>
    public sealed class AmountConversionTransform : IRangeTransform
    {
        private readonly AmountConversionOptions _options;

        public AmountConversionTransform(AmountConversionOptions options)
        {
            _options = options;
        }

        public string Id
        {
            get { return "amountConversion"; }
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

            if (cell.IsFormula)
            {
                var formula = cell.Formula;
                if (formula == null || !formula.CanTransform)
                {
                    return CellTransformOutcome.Skip(cell, SkipReason.FormulaNotTransformable);
                }

                // 公式保持公式属性：原公式 ×/÷ Rate，再套 ROUND。
                var transformed = AmountConversionService.BuildFormula(formula.Expression, _options);
                return CellTransformOutcome.WriteFormula(cell, transformed);
            }

            if (cell.CellType == CellValueType.ConstantNumber)
            {
                if (!cell.TryGetDecimal(out var value))
                {
                    return CellTransformOutcome.Skip(cell, SkipReason.UnsupportedType);
                }

                var result = AmountConversionService.ConvertToResult(value, _options);
                return CellTransformOutcome.WriteValue(cell, result);
            }

            return CellTransformOutcome.Skip(cell, Rounding.RoundingTransform.MapSkipReason(cell.CellType));
        }
    }
}
