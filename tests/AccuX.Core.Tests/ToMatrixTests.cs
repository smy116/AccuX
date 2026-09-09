using System;
using AccuX.Host;
using Xunit;

namespace AccuX.Core.Tests
{
    /// <summary>
    /// COM 返回值归一化测试（规格 §5.2 / §16）。
    /// <para>
    /// 回归重点：Excel/WPS 的 COM 二维数组下界为 1，若直接透传，调用方按 0 基索引访问会抛
    /// IndexOutOfRangeException（曾导致所有功能报“索引超出了数组界限”）。
    /// </para>
    /// </summary>
    public class ToMatrixTests
    {
        private static Array CreateOneBased(int rows, int columns)
        {
            var array = Array.CreateInstance(typeof(object), new[] { rows, columns }, new[] { 1, 1 });
            for (var row = 1; row <= rows; row++)
            {
                for (var column = 1; column <= columns; column++)
                {
                    array.SetValue((double)(row * 10 + column), row, column);
                }
            }

            return array;
        }

        [Fact]
        public void ToMatrix_OneBasedSameSize_ReturnsZeroBasedCopy()
        {
            var raw = CreateOneBased(2, 2);

            var matrix = ExcelRangeOperationHost.ToMatrix(raw, 2, 2);

            Assert.Equal(0, matrix.GetLowerBound(0));
            Assert.Equal(0, matrix.GetLowerBound(1));
            Assert.Equal(11d, matrix[0, 0]);
            Assert.Equal(12d, matrix[0, 1]);
            Assert.Equal(21d, matrix[1, 0]);
            Assert.Equal(22d, matrix[1, 1]);
        }

        [Fact]
        public void ToMatrix_OneBasedDifferentSize_CopiesOverlap()
        {
            var raw = CreateOneBased(2, 3);

            var matrix = ExcelRangeOperationHost.ToMatrix(raw, 2, 2);

            Assert.Equal(11d, matrix[0, 0]);
            Assert.Equal(12d, matrix[0, 1]);
            Assert.Equal(21d, matrix[1, 0]);
            Assert.Equal(22d, matrix[1, 1]);
        }

        [Fact]
        public void ToMatrix_ZeroBased_IsAlsoSupported()
        {
            var raw = new object[1, 2];
            raw[0, 0] = "a";
            raw[0, 1] = "b";

            var matrix = ExcelRangeOperationHost.ToMatrix(raw, 1, 2);

            Assert.Equal("a", matrix[0, 0]);
            Assert.Equal("b", matrix[0, 1]);
        }

        [Fact]
        public void ToMatrix_Scalar_BroadcastsToTargetSize()
        {
            var matrix = ExcelRangeOperationHost.ToMatrix(7d, 2, 2);

            Assert.Equal(7d, matrix[0, 0]);
            Assert.Equal(7d, matrix[0, 1]);
            Assert.Equal(7d, matrix[1, 0]);
            Assert.Equal(7d, matrix[1, 1]);
        }

        [Fact]
        public void ToMatrix_NullScalar_BroadcastsNulls()
        {
            var matrix = ExcelRangeOperationHost.ToMatrix(null, 1, 2);

            Assert.Null(matrix[0, 0]);
            Assert.Null(matrix[0, 1]);
        }

        [Fact]
        public void ToMatrix_LargerTargetThanSource_LeavesRestNull()
        {
            var raw = CreateOneBased(1, 1);

            var matrix = ExcelRangeOperationHost.ToMatrix(raw, 2, 2);

            Assert.Equal(11d, matrix[0, 0]);
            Assert.Null(matrix[0, 1]);
            Assert.Null(matrix[1, 0]);
            Assert.Null(matrix[1, 1]);
        }
    }
}
