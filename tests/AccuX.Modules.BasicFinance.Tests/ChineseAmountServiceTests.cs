using System;
using AccuX.Modules.BasicFinance.ChineseAmount;
using Xunit;

namespace AccuX.Modules.BasicFinance.Tests
{
    public class ChineseAmountServiceTests
    {
        [Theory]
        [InlineData(0, "零元整")]
        [InlineData(1, "壹元整")]
        [InlineData(10, "壹拾元整")]
        [InlineData(100, "壹佰元整")]
        [InlineData(1000, "壹仟元整")]
        [InlineData(10000, "壹万元整")]
        [InlineData(100000000, "壹亿元整")]
        [InlineData(123456789, "壹亿贰仟叁佰肆拾伍万陆仟柒佰捌拾玖元整")]
        public void IntegerAmounts(decimal amount, string expected)
        {
            Assert.Equal(expected, ChineseAmountService.ToChineseAmount(amount));
        }

        [Theory]
        [InlineData(0.5, "零元伍角")]
        [InlineData(0.05, "零元零伍分")]
        [InlineData(0.55, "零元伍角伍分")]
        [InlineData(1.01, "壹元零壹分")]
        [InlineData(1.10, "壹元壹角")]
        [InlineData(1.11, "壹元壹角壹分")]
        [InlineData(1.00, "壹元整")]
        public void DecimalAmounts(decimal amount, string expected)
        {
            Assert.Equal(expected, ChineseAmountService.ToChineseAmount(amount));
        }

        [Theory]
        [InlineData(-1, "负壹元整")]
        [InlineData(-1.5, "负壹元伍角")]
        [InlineData(-12345.67, "负壹万贰仟叁佰肆拾伍元陆角柒分")]
        public void NegativeAmounts(decimal amount, string expected)
        {
            Assert.Equal(expected, ChineseAmountService.ToChineseAmount(amount));
        }

        [Theory]
        [InlineData(10001, "壹万零壹元整")]
        [InlineData(10010, "壹万零壹拾元整")]
        [InlineData(1000000, "壹佰万元整")]
        [InlineData(1000000000000, "壹万亿元整")]
        public void ZeroHandlingBetweenGroups(decimal amount, string expected)
        {
            Assert.Equal(expected, ChineseAmountService.ToChineseAmount(amount));
        }

        [Fact]
        public void RoundsToCents()
        {
            // 0.005 舍入到分 -> 0.01
            Assert.Equal("零元零壹分", ChineseAmountService.ToChineseAmount(0.005m));
        }

        [Fact]
        public void OutOfRange_Throws()
        {
            Assert.Throws<ArgumentOutOfRangeException>(
                () => ChineseAmountService.ToChineseAmount(ChineseAmountService.MaxAmount + 1m));
        }

        [Fact]
        public void TryToChineseAmount_ReportsFailure()
        {
            Assert.False(ChineseAmountService.TryToChineseAmount(
                ChineseAmountService.MaxAmount + 1m, out _, out var error));
            Assert.False(string.IsNullOrEmpty(error));
        }
    }
}
