using AccuX.Core.Operations;
using AccuX.Host;
using Xunit;

namespace AccuX.Core.Tests
{
    public class MarkHostColorTests
    {
        [Theory]
        [InlineData("#18be6a", 0x6abe18)]
        [InlineData("#ed4015", 0x1540ed)]
        [InlineData("#fe9900", 0x0099fe)]
        [InlineData("#2db7f5", 0xf5b72d)]
        public void ParseOleColor_ConvertsHtmlRgbToExcelOleColor(string hexColor, int expected)
        {
            Assert.Equal(expected, ExcelRangeOperationHost.ParseOleColor(hexColor));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("18be6a")]
        [InlineData("#18be6")]
        [InlineData("#ggbe6a")]
        public void ParseOleColor_RejectsInvalidColors(string hexColor)
        {
            var exception = Assert.Throws<HostOperationException>(
                () => ExcelRangeOperationHost.ParseOleColor(hexColor));

            Assert.Contains("#RRGGBB", exception.Message);
        }
    }
}
