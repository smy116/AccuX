using AccuX.Core.Cells;
using Xunit;

namespace AccuX.Core.Tests
{
    /// <summary>
    /// CellValueClassifier 是 V1 唯一的分类入口，必须有独立测试（规格 §26 Core Unit Test）。
    /// </summary>
    public class CellValueClassifierTests
    {
        [Fact]
        public void Blank_WhenNoValue()
        {
            var type = CellValueClassifier.Classify(new CellClassificationInput { IsBlank = true });
            Assert.Equal(CellValueType.Blank, type);
        }

        [Fact]
        public void ConstantNumber_ForNumericValue()
        {
            var type = CellValueClassifier.Classify(new CellClassificationInput { Value = 12.34m });
            Assert.Equal(CellValueType.ConstantNumber, type);
        }

        [Fact]
        public void FormulaNumber_ForNumericFormula()
        {
            var type = CellValueClassifier.Classify(new CellClassificationInput { HasFormula = true, Value = 12.34m });
            Assert.Equal(CellValueType.FormulaNumber, type);
        }

        [Fact]
        public void Text_ForStringValue()
        {
            var type = CellValueClassifier.Classify(new CellClassificationInput { Value = "abc" });
            Assert.Equal(CellValueType.Text, type);
        }

        [Fact]
        public void Date_WhenHostFlagsDate()
        {
            var type = CellValueClassifier.Classify(new CellClassificationInput { Value = 45000m, IsDate = true });
            Assert.Equal(CellValueType.Date, type);
        }

        [Fact]
        public void Date_WhenValueIsDateTime()
        {
            var type = CellValueClassifier.Classify(new CellClassificationInput { Value = System.DateTime.Today });
            Assert.Equal(CellValueType.Date, type);
        }

        [Fact]
        public void Boolean_ForBooleanValue()
        {
            var type = CellValueClassifier.Classify(new CellClassificationInput { Value = true });
            Assert.Equal(CellValueType.Boolean, type);
        }

        [Fact]
        public void Error_WhenHostFlagsError()
        {
            var type = CellValueClassifier.Classify(new CellClassificationInput { Value = -2146826281, IsError = true });
            Assert.Equal(CellValueType.Error, type);
        }

        [Fact]
        public void FormulaError_ForErrorFormula()
        {
            var type = CellValueClassifier.Classify(new CellClassificationInput
            {
                HasFormula = true,
                Value = -2146826281,
                IsError = true
            });
            Assert.Equal(CellValueType.FormulaError, type);
        }

        [Fact]
        public void FormulaText_ForTextFormula()
        {
            var type = CellValueClassifier.Classify(new CellClassificationInput { HasFormula = true, Value = "abc" });
            Assert.Equal(CellValueType.FormulaText, type);
        }

        [Fact]
        public void FormulaDate_ForDateFormula()
        {
            var type = CellValueClassifier.Classify(new CellClassificationInput
            {
                HasFormula = true,
                Value = 45000m,
                IsDate = true
            });
            Assert.Equal(CellValueType.FormulaDate, type);
        }

        [Fact]
        public void FormulaBoolean_ForBooleanFormula()
        {
            var type = CellValueClassifier.Classify(new CellClassificationInput { HasFormula = true, Value = false });
            Assert.Equal(CellValueType.FormulaBoolean, type);
        }

        [Theory]
        [InlineData(1)]
        [InlineData(1.5)]
        [InlineData(1L)]
        [InlineData((short)1)]
        public void NormalizeNumeric_ConvertsToDecimal(object raw)
        {
            var normalized = CellValueClassifier.NormalizeNumeric(raw);
            Assert.IsType<decimal>(normalized);
        }

        [Fact]
        public void IsNumber_ExtensionWorks()
        {
            Assert.True(CellValueType.ConstantNumber.IsNumber());
            Assert.True(CellValueType.FormulaNumber.IsNumber());
            Assert.False(CellValueType.Text.IsNumber());
        }

        [Fact]
        public void IsFormula_ExtensionWorks()
        {
            Assert.True(CellValueType.FormulaNumber.IsFormula());
            Assert.True(CellValueType.FormulaDate.IsFormula());
            Assert.False(CellValueType.Date.IsFormula());
            Assert.False(CellValueType.ConstantNumber.IsFormula());
        }
    }
}
