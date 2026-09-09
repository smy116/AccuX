using System;
using AccuX.Modules.BasicFinance.Rounding;
using Xunit;

namespace AccuX.Modules.BasicFinance.Tests
{
    public class RoundingServiceTests
    {
        [Theory]
        [InlineData(1.005, 2, 1.01)]
        [InlineData(1.004, 2, 1.00)]
        [InlineData(2.675, 2, 2.68)]
        [InlineData(1.5, 0, 2)]
        [InlineData(2.5, 0, 3)]
        [InlineData(-1.5, 0, -2)]
        [InlineData(-2.5, 0, -3)]
        [InlineData(1.2345, 3, 1.235)]
        public void Round_UsesAwayFromZero(decimal value, int digits, decimal expected)
        {
            Assert.Equal(expected, RoundingService.Round(value, digits));
        }

        [Fact]
        public void Round_Midpoint_IsNotBankersRounding()
        {
            // 银行家舍入会把 2.5 变成 2，财务规则必须是 3。
            Assert.Equal(3m, RoundingService.Round(2.5m, 0));
        }

        [Fact]
        public void WouldChange_DetectsDifference()
        {
            Assert.True(RoundingService.WouldChange(1.005m, 2));
            Assert.False(RoundingService.WouldChange(1.00m, 2));
        }

        [Theory]
        [InlineData(-1)]
        [InlineData(16)]
        public void Round_OutOfRange_Throws(int digits)
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => RoundingService.Round(1m, digits));
        }

        [Theory]
        [InlineData(0, true)]
        [InlineData(15, true)]
        [InlineData(-1, false)]
        [InlineData(16, false)]
        public void Options_Validation(int digits, bool expected)
        {
            Assert.Equal(expected, RoundingOptions.TryNormalize(digits, out var options, out _));
            if (expected)
            {
                Assert.Equal(digits, options.Digits);
            }
        }
    }
}
