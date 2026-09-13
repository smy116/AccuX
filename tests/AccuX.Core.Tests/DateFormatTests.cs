using AccuX.Host;
using Xunit;

namespace AccuX.Core.Tests
{
    public class DateFormatTests
    {
        [Theory]
        [InlineData("#,##0.00;[Red]-#,##0.00", false)]
        [InlineData("0.00\"days\"", false)]
        [InlineData("0.00\\d", false)]
        [InlineData("0.00_d", false)]
        [InlineData("0.00m", false)]
        [InlineData("yyyy-mm-dd", true)]
        [InlineData("m/d/yy h:mm", true)]
        [InlineData("mm", true)]
        [InlineData("[h]:mm:ss", true)]
        public void IsDateFormat_IgnoresLiteralsColorsAndConditions(
            string numberFormat,
            bool expected)
        {
            Assert.Equal(expected, ExcelRangeOperationHost.IsDateFormat(numberFormat));
        }
    }
}
