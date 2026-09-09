using AccuX.Modules.BasicFinance.AmountConversion;
using Xunit;

namespace AccuX.Modules.BasicFinance.Tests
{
    public class AmountConversionServiceTests
    {
        [Fact]
        public void DivideByTenThousand()
        {
            var options = new AmountConversionOptions { Mode = ConversionMode.Divide, Rate = 10000m, Digits = 2 };
            Assert.Equal(1.23m, AmountConversionService.Convert(12345m, options));
        }

        [Fact]
        public void DivideByHundred()
        {
            var options = new AmountConversionOptions { Mode = ConversionMode.Divide, Rate = 100m, Digits = 2 };
            Assert.Equal(123.45m, AmountConversionService.Convert(12345m, options));
        }

        [Fact]
        public void Multiply()
        {
            var options = new AmountConversionOptions { Mode = ConversionMode.Multiply, Rate = 10000m, Digits = 2 };
            Assert.Equal(123450000m, AmountConversionService.Convert(12345m, options));
        }

        [Fact]
        public void Rounding_AppliedAfterConversion()
        {
            var options = new AmountConversionOptions { Mode = ConversionMode.Divide, Rate = 3m, Digits = 2 };
            Assert.Equal(1.00m, AmountConversionService.Convert(3m, options));
        }

        [Fact]
        public void DivideByZero_Throws()
        {
            var options = new AmountConversionOptions { Mode = ConversionMode.Divide, Rate = 0m };
            Assert.Throws<System.DivideByZeroException>(() => AmountConversionService.Convert(100m, options));
        }

        [Fact]
        public void Validate_DivideByZero_IsRejected()
        {
            var options = new AmountConversionOptions { Mode = ConversionMode.Divide, Rate = 0m };
            Assert.False(AmountConversionOptions.TryValidate(options, out var error));
            Assert.Contains("0", error);
        }

        [Fact]
        public void AppendWanSuffix_ProducesText()
        {
            var options = new AmountConversionOptions
            {
                Mode = ConversionMode.Divide,
                Rate = 10000m,
                Digits = 2,
                AppendWanSuffix = true
            };

            var result = AmountConversionService.ConvertToResult(12345m, options);

            Assert.Equal("1.23万", result);
        }

        [Fact]
        public void NoSuffix_ProducesDecimal()
        {
            var options = new AmountConversionOptions { Mode = ConversionMode.Divide, Rate = 10000m, Digits = 2 };
            var result = AmountConversionService.ConvertToResult(12345m, options);

            Assert.IsType<decimal>(result);
        }

        [Fact]
        public void BuildFormula_KeepsFormulaAndAddsRound()
        {
            var options = new AmountConversionOptions { Mode = ConversionMode.Divide, Rate = 10000m, Digits = 2 };
            var formula = AmountConversionService.BuildFormula("=A1", options);

            Assert.Equal("=ROUND((A1)/10000,2)", formula);
        }

        [Fact]
        public void BuildFormula_MultiplyMode()
        {
            var options = new AmountConversionOptions { Mode = ConversionMode.Multiply, Rate = 100m, Digits = 0 };
            var formula = AmountConversionService.BuildFormula("=SUM(A1:A3)", options);

            Assert.Equal("=ROUND((SUM(A1:A3))*100,0)", formula);
        }
    }
}
