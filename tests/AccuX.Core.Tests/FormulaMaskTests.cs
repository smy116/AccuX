using AccuX.Host;
using Xunit;

namespace AccuX.Core.Tests
{
    /// <summary>
    /// 公式区域掩码构建测试（规格 §8 Formula 处理契约）。
    /// SpecialCells 返回的区域是绝对地址，需要按选区基准映射回 0 基掩码。
    /// </summary>
    public class FormulaMaskTests
    {
        [Fact]
        public void MarkArea_SingleCell_MarksOnlyThatCell()
        {
            var mask = new bool[3, 1];

            ExcelRangeOperationHost.MarkArea(mask, 1, 0, 1, 1);

            Assert.False(mask[0, 0]);
            Assert.True(mask[1, 0]);
            Assert.False(mask[2, 0]);
        }

        [Fact]
        public void MarkArea_MultipleAreas_Accumulate()
        {
            var mask = new bool[3, 1];

            ExcelRangeOperationHost.MarkArea(mask, 0, 0, 1, 1);
            ExcelRangeOperationHost.MarkArea(mask, 2, 0, 1, 1);

            Assert.True(mask[0, 0]);
            Assert.False(mask[1, 0]);
            Assert.True(mask[2, 0]);
        }

        [Fact]
        public void MarkArea_OutOfBounds_IsClipped()
        {
            var mask = new bool[2, 2];

            ExcelRangeOperationHost.MarkArea(mask, 1, 1, 5, 5);

            Assert.False(mask[0, 0]);
            Assert.True(mask[1, 1]);
        }

        [Fact]
        public void MarkArea_NegativeStart_IsIgnored()
        {
            var mask = new bool[2, 2];

            ExcelRangeOperationHost.MarkArea(mask, -1, -1, 2, 2);

            Assert.True(mask[0, 0]);
            Assert.False(mask[1, 1]);
        }

        [Fact]
        public void MarkArea_NullMask_DoesNotThrow()
        {
            ExcelRangeOperationHost.MarkArea(null, 0, 0, 1, 1);
        }
    }
}
