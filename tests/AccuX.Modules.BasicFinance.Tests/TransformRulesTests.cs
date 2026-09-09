using AccuX.Core.Cells;
using AccuX.Core.Operations;
using AccuX.Modules.BasicFinance.AmountConversion;
using AccuX.Modules.BasicFinance.ChineseAmount;
using AccuX.Modules.BasicFinance.Rounding;
using Xunit;

namespace AccuX.Modules.BasicFinance.Tests
{
    /// <summary>
    /// 验证规格 §15 的数据处理规则表在修改型功能中一致生效。
    /// </summary>
    public class TransformRulesTests
    {
        private static CellData Cell(CellValueType type, object value, FormulaInfo formula = null)
        {
            return new CellData(0, 0)
            {
                CellType = type,
                Value = value,
                Formula = formula,
                Writable = true
            };
        }

        private static readonly OperationContext Context =
            new OperationContext("cmd", "mod", new RangeTarget("b", "s", "Sheet1", "A1", 1, 1, 1, false, false));

        [Theory]
        [InlineData(CellValueType.Text, SkipReason.Text)]
        [InlineData(CellValueType.Date, SkipReason.Date)]
        [InlineData(CellValueType.Boolean, SkipReason.Boolean)]
        [InlineData(CellValueType.Blank, SkipReason.Blank)]
        [InlineData(CellValueType.Error, SkipReason.Error)]
        [InlineData(CellValueType.Unsupported, SkipReason.UnsupportedType)]
        public void AllTransforms_SkipNonAmountTypes(CellValueType type, SkipReason expected)
        {
            var cell = Cell(type, null);

            AssertSkip(new RoundingTransform(2), cell, expected);
            AssertSkip(new AmountConversionTransform(new AmountConversionOptions()), cell, expected);
            AssertSkip(new ChineseAmountTransform(), cell, expected);
        }

        [Fact]
        public void Rounding_Number_ProducesValueWrite()
        {
            var outcome = new RoundingTransform(2).Transform(Cell(CellValueType.ConstantNumber, 1.005m), Context);

            Assert.True(outcome.HasWrite);
            Assert.Equal(1.01m, outcome.Value);
        }

        [Fact]
        public void Rounding_Formula_ProducesFormulaWrite()
        {
            var cell = Cell(CellValueType.FormulaNumber, 1m, new FormulaInfo("=A1", FormulaKind.Normal, true));
            var outcome = new RoundingTransform(2).Transform(cell, Context);

            Assert.True(outcome.IsFormulaWrite);
            Assert.Equal("=ROUND(A1,2)", outcome.Formula);
        }

        [Fact]
        public void Rounding_UntransformableFormula_Skipped()
        {
            var cell = Cell(CellValueType.FormulaNumber, 1m, new FormulaInfo("={A1}", FormulaKind.Array, false));
            var outcome = new RoundingTransform(2).Transform(cell, Context);

            Assert.False(outcome.HasWrite);
            Assert.Equal(SkipReason.FormulaNotTransformable, outcome.SkipReason);
        }

        [Fact]
        public void ChineseAmount_Number_ProducesTextWrite()
        {
            var outcome = new ChineseAmountTransform().Transform(Cell(CellValueType.ConstantNumber, 1234.56m), Context);

            Assert.True(outcome.HasWrite);
            Assert.Equal("壹仟贰佰叁拾肆元伍角陆分", outcome.Value);
        }

        [Fact]
        public void ChineseAmount_Formula_ProducesTextWriteNotFormula()
        {
            var cell = Cell(CellValueType.FormulaNumber, 1m, new FormulaInfo("=A1", FormulaKind.Normal, true));
            var outcome = new ChineseAmountTransform().Transform(cell, Context);

            Assert.True(outcome.HasWrite);
            Assert.False(outcome.IsFormulaWrite);
            Assert.Equal("壹元整", outcome.Value);
        }

        [Fact]
        public void AmountConversion_Formula_KeepsFormula()
        {
            var cell = Cell(CellValueType.FormulaNumber, 1m, new FormulaInfo("=A1", FormulaKind.Normal, true));
            var options = new AmountConversionOptions { Mode = ConversionMode.Divide, Rate = 10000m, Digits = 2 };
            var outcome = new AmountConversionTransform(options).Transform(cell, Context);

            Assert.True(outcome.IsFormulaWrite);
            Assert.Equal("=ROUND((A1)/10000,2)", outcome.Formula);
        }

        [Fact]
        public void NotWritable_IsSkipped()
        {
            var cell = Cell(CellValueType.ConstantNumber, 1m);
            cell.Writable = false;

            var outcome = new RoundingTransform(2).Transform(cell, Context);

            Assert.False(outcome.HasWrite);
            Assert.Equal(SkipReason.NotWritable, outcome.SkipReason);
        }

        private static void AssertSkip(IRangeTransform transform, CellData cell, SkipReason expected)
        {
            var outcome = transform.Transform(cell, Context);
            Assert.NotNull(outcome);
            Assert.False(outcome.HasWrite);
            Assert.Equal(expected, outcome.SkipReason);
        }
    }
}
