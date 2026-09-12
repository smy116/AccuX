using AccuX.Host;
using Xunit;

namespace AccuX.Core.Tests
{
    /// <summary>
    /// UsedRange 选区裁剪的纯逻辑测试，不依赖 Excel/WPS COM 实例。
    /// </summary>
    public sealed class UsedRangeSelectionTests
    {
        [Fact]
        public void TryIntersectRangeBounds_WhenSelectionIsInside_ReturnsSelectionBounds()
        {
            var selection = Bounds(3, 10, 4, 8);
            var usedRange = Bounds(1, 20, 1, 20);

            var success = ExcelHostBase.TryIntersectRangeBounds(selection, usedRange, out var intersection);

            Assert.True(success);
            Assert.Equal(selection, intersection);
        }

        [Fact]
        public void TryIntersectRangeBounds_WhenSelectionPartiallyOverlaps_ClipsAllFourEdges()
        {
            var selection = Bounds(1, 10, 2, 12);
            var usedRange = Bounds(4, 8, 6, 10);

            var success = ExcelHostBase.TryIntersectRangeBounds(selection, usedRange, out var intersection);

            Assert.True(success);
            Assert.Equal(Bounds(4, 8, 6, 10), intersection);
        }

        [Fact]
        public void TryIntersectRangeBounds_WhenSelectionContainsUsedRange_ReturnsUsedRangeBounds()
        {
            var selection = Bounds(4, 20, 3, 20);
            var usedRange = Bounds(8, 12, 7, 14);

            var success = ExcelHostBase.TryIntersectRangeBounds(selection, usedRange, out var intersection);

            Assert.True(success);
            Assert.Equal(usedRange, intersection);
        }

        [Theory]
        [InlineData(1, 3, 1, 3, 4, 8, 1, 3)]
        [InlineData(1, 3, 1, 3, 1, 3, 4, 8)]
        public void TryIntersectRangeBounds_WhenRangesOnlyTouchAtAnEdge_ReturnsNoIntersection(
            int selectionFirstRow,
            int selectionLastRow,
            int selectionFirstColumn,
            int selectionLastColumn,
            int usedFirstRow,
            int usedLastRow,
            int usedFirstColumn,
            int usedLastColumn)
        {
            var success = ExcelHostBase.TryIntersectRangeBounds(
                Bounds(selectionFirstRow, selectionLastRow, selectionFirstColumn, selectionLastColumn),
                Bounds(usedFirstRow, usedLastRow, usedFirstColumn, usedLastColumn),
                out var intersection);

            Assert.False(success);
            Assert.Equal(default(ExcelHostBase.RangeBounds), intersection);
        }

        [Fact]
        public void TryIntersectRangeBounds_ReturnsBoundsWhoseCellCountMatchesClippedTarget()
        {
            var success = ExcelHostBase.TryIntersectRangeBounds(
                Bounds(2, 12, 3, 15),
                Bounds(5, 8, 7, 10),
                out var intersection);

            Assert.True(success);
            var rows = intersection.LastRow - intersection.FirstRow + 1;
            var columns = intersection.LastColumn - intersection.FirstColumn + 1;
            Assert.Equal(4, rows);
            Assert.Equal(4, columns);
            Assert.Equal(16, rows * columns);
        }

        private static ExcelHostBase.RangeBounds Bounds(
            int firstRow,
            int lastRow,
            int firstColumn,
            int lastColumn)
        {
            return new ExcelHostBase.RangeBounds(firstRow, lastRow, firstColumn, lastColumn);
        }
    }
}
