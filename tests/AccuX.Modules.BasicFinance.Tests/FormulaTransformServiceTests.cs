using AccuX.Modules.BasicFinance.Common;
using Xunit;

namespace AccuX.Modules.BasicFinance.Tests
{
    public class FormulaTransformServiceTests
    {
        [Theory]
        [InlineData("=A1", "A1")]
        [InlineData("A1", "A1")]
        [InlineData("  =SUM(A1:A2)  ", "SUM(A1:A2)")]
        public void GetBody_StripsLeadingEquals(string input, string expected)
        {
            Assert.Equal(expected, FormulaTransformService.GetBody(input));
        }

        [Theory]
        [InlineData("A1", "=A1")]
        [InlineData("=A1", "=A1")]
        public void EnsureLeadingEquals(string input, string expected)
        {
            Assert.Equal(expected, FormulaTransformService.EnsureLeadingEquals(input));
        }

        [Fact]
        public void WrapFunction_BuildsRound()
        {
            Assert.Equal("=ROUND(A1,2)", FormulaTransformService.WrapFunction("=A1", "ROUND", "2"));
        }

        [Fact]
        public void ApplyArithmetic_Multiply()
        {
            Assert.Equal("=(A1)*7", FormulaTransformService.ApplyArithmetic("=A1", "*", 7m));
        }

        [Fact]
        public void ApplyArithmetic_Divide()
        {
            Assert.Equal("=(A1)/10000", FormulaTransformService.ApplyArithmetic("=A1", "/", 10000m));
        }

        [Fact]
        public void FormatLiteral_TrimsTrailingZeros()
        {
            Assert.Equal("7", FormulaTransformService.FormatLiteral(7.00m));
            Assert.Equal("7.5", FormulaTransformService.FormatLiteral(7.50m));
            Assert.Equal("0.01", FormulaTransformService.FormatLiteral(0.0100m));
        }

        [Fact]
        public void RepeatedApplication_IsNotDeduplicated()
        {
            // 规格 §17.2：连续两次显式业务操作是合法的，Service 不去重。
            var once = FormulaTransformService.ApplyArithmetic("=A1", "*", 7m);
            var twice = FormulaTransformService.ApplyArithmetic(once, "*", 7m);

            Assert.Equal("=(A1)*7", once);
            Assert.Equal("=((A1)*7)*7", twice);
        }

        [Fact]
        public void RepeatedRounding_WrapsAgain()
        {
            var once = FormulaTransformService.WrapFunction("=A1", "ROUND", "2");
            var twice = FormulaTransformService.WrapFunction(once, "ROUND", "2");

            Assert.Equal("=ROUND(ROUND(A1,2),2)", twice);
        }
    }
}
