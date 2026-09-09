using AccuX.Core.Cells;
using AccuX.Core.Operations;

namespace AccuX.Modules.BasicFinance.ChineseAmount
{
    /// <summary>
    /// 金额大写的业务转换（规格 §14 / §15）。
    /// 普通数值与公式单元格都直接转换为中文大写文字，原地覆盖选区（§14.3）。
    /// </summary>
    public sealed class ChineseAmountTransform : IRangeTransform
    {
        public string Id
        {
            get { return "chineseAmount"; }
        }

        public bool MayWriteFormula
        {
            get { return false; }
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

            // 普通数值与公式（含复杂公式）都按其计算结果输出大写文字。
            if (cell.CellType != CellValueType.ConstantNumber && cell.CellType != CellValueType.FormulaNumber)
            {
                return CellTransformOutcome.Skip(cell, Rounding.RoundingTransform.MapSkipReason(cell.CellType));
            }

            if (!cell.TryGetDecimal(out var value))
            {
                return CellTransformOutcome.Skip(cell, SkipReason.UnsupportedType);
            }

            if (!ChineseAmountService.TryToChineseAmount(value, out var text, out _))
            {
                return CellTransformOutcome.Skip(cell, SkipReason.UnsupportedType);
            }

            // 文本结果不设置 NumberFormat，避免宿主按数值格式解释。
            return CellTransformOutcome.WriteValue(cell, text, "@");
        }
    }
}
