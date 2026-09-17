using System;
using System.Collections.Generic;
using AccuX.Core.Operations;
using Excel = Microsoft.Office.Interop.Excel;

namespace AccuX.Host
{
    internal static class ExcelSelectionAreas
    {
        internal const int MaxAreas = 4096;

        internal static IReadOnlyList<RangeArea> Capture(
            Excel.Worksheet worksheet, Excel.Range selection, long maxCells)
        {
            try
            {
                if (worksheet == null || selection == null)
                {
                    throw new HostOperationException("无法读取当前工作表或选区，请重新选择。");
                }

                // 不使用 SafeAreaCount：读取失败绝不能当成单区域。
                var areas = selection.Areas;
                if (areas == null)
                {
                    throw new HostOperationException("无法确定当前选区的区域数量。");
                }

                var count = areas.Count;
                ValidateAreaCount(count);
                var usedBounds = ReadBounds(worksheet.UsedRange);
                var input = new List<Bounds>(count);
                for (var index = 1; index <= count; index++)
                {
                    input.Add(ReadBounds(areas[index]));
                }

                var normalized = Normalize(input, usedBounds, maxCells);
                var result = new List<RangeArea>(normalized.Count);
                foreach (var bounds in normalized)
                {
                    var firstCell = (Excel.Range)worksheet.Cells[bounds.FirstRow, bounds.FirstColumn];
                    var rectangle = firstCell.Resize[bounds.RowCount, bounds.ColumnCount];
                    if (rectangle == null)
                    {
                        throw new HostOperationException("无法构造选区的单矩形区域。");
                    }

                    EnsureUnmerged(rectangle.MergeCells);
                    var address = rectangle.get_Address(true, true, Excel.XlReferenceStyle.xlA1,
                        false, Type.Missing);
                    if (string.IsNullOrWhiteSpace(address))
                    {
                        throw new HostOperationException("无法读取选区的单矩形地址。");
                    }

                    result.Add(new RangeArea(address, bounds.RowCount, bounds.ColumnCount));
                }

                return result.AsReadOnly();
            }
            catch (HostOperationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new HostOperationException("无法读取或规范化当前选区，请重新选择后重试。", ex);
            }
        }

        internal static void ValidateAreaCount(int count)
        {
            if (count <= 0)
            {
                throw new HostOperationException("无法确定当前选区的有效区域数量。");
            }

            if (count > MaxAreas)
            {
                throw TooManyAreas();
            }
        }

        internal static void EnsureUnmerged(object mergeCells)
        {
            if (!(mergeCells is bool merged) || merged)
            {
                throw new HostOperationException("不支持合并单元格，或无法确认选区是否包含合并单元格。");
            }
        }

        private static Bounds ReadBounds(Excel.Range range)
        {
            if (range == null)
            {
                throw new HostOperationException("无法确定选区或 UsedRange 的矩形边界。");
            }

            var firstRow = range.Row;
            var firstColumn = range.Column;
            var rows = range.Rows.Count;
            var columns = range.Columns.Count;
            if (rows <= 0 || columns <= 0)
            {
                throw new HostOperationException("选区或 UsedRange 的矩形大小无效。");
            }

            return new Bounds(firstRow, checked(firstRow + (rows - 1)),
                firstColumn, checked(firstColumn + (columns - 1)));
        }

        /// <summary>
        /// 按输入顺序裁剪、扣除已接纳的矩形。仅操作矩形，不枚举单元格。
        /// 输入数量和处理中间碎片均有硬上限，防止碎片爆炸。
        /// </summary>
        internal static IReadOnlyList<Bounds> Normalize(
            IEnumerable<Bounds> areas, Bounds usedRange, long maxCells)
        {
            if (areas == null)
            {
                throw new HostOperationException("无法读取选区区域列表。");
            }

            usedRange.Validate();
            if (maxCells < 0)
            {
                throw new HostOperationException("选区单元格数量上限不能为负数。");
            }

            var result = new List<Bounds>();
            long total = 0;
            var inputCount = 0;
            foreach (var area in areas)
            {
                ValidateAreaCount(++inputCount);
                area.Validate();
                if (!TryIntersect(area, usedRange, out var clipped))
                {
                    continue;
                }

                var pending = new List<Bounds> { clipped };
                foreach (var accepted in result)
                {
                    var next = new List<Bounds>();
                    foreach (var fragment in pending)
                    {
                        foreach (var remainder in Subtract(fragment, accepted))
                        {
                            if (result.Count + next.Count >= MaxAreas)
                            {
                                throw TooManyAreas();
                            }

                            next.Add(remainder);
                        }
                    }

                    pending = next;
                    if (pending.Count == 0)
                    {
                        break;
                    }
                }

                foreach (var fragment in pending)
                {
                    if (result.Count >= MaxAreas)
                    {
                        throw TooManyAreas();
                    }

                    // 用减法比较避免 long 累加溢出，且只计入已去重的单元格。
                    if (fragment.CellCount > maxCells - total)
                    {
                        throw new HostOperationException("选区去重后的单元格数量超过上限（" + maxCells + "）。");
                    }

                    total += fragment.CellCount;
                    result.Add(fragment);
                }
            }

            if (result.Count == 0)
            {
                throw new HostOperationException("当前选区在 UsedRange 内没有单元格，请重新选择。");
            }

            return result.AsReadOnly();
        }

        /// <summary>source 减去 cut，依次输出上、下、左、右，最多四个不相交矩形。</summary>
        internal static IReadOnlyList<Bounds> Subtract(Bounds source, Bounds cut)
        {
            source.Validate();
            cut.Validate();
            var result = new List<Bounds>(4);
            if (!TryIntersect(source, cut, out var overlap))
            {
                result.Add(source);
                return result;
            }

            if (source.FirstRow < overlap.FirstRow)
                result.Add(new Bounds(source.FirstRow, overlap.FirstRow - 1, source.FirstColumn, source.LastColumn));
            if (overlap.LastRow < source.LastRow)
                result.Add(new Bounds(overlap.LastRow + 1, source.LastRow, source.FirstColumn, source.LastColumn));
            if (source.FirstColumn < overlap.FirstColumn)
                result.Add(new Bounds(overlap.FirstRow, overlap.LastRow, source.FirstColumn, overlap.FirstColumn - 1));
            if (overlap.LastColumn < source.LastColumn)
                result.Add(new Bounds(overlap.FirstRow, overlap.LastRow, overlap.LastColumn + 1, source.LastColumn));
            return result;
        }

        private static bool TryIntersect(Bounds left, Bounds right, out Bounds intersection)
        {
            var firstRow = Math.Max(left.FirstRow, right.FirstRow);
            var lastRow = Math.Min(left.LastRow, right.LastRow);
            var firstColumn = Math.Max(left.FirstColumn, right.FirstColumn);
            var lastColumn = Math.Min(left.LastColumn, right.LastColumn);
            if (firstRow > lastRow || firstColumn > lastColumn)
            {
                intersection = default(Bounds);
                return false;
            }

            intersection = new Bounds(firstRow, lastRow, firstColumn, lastColumn);
            return true;
        }

        private static HostOperationException TooManyAreas()
        {
            return new HostOperationException("选区区域数量或去重后的碎片数量超过上限（" + MaxAreas + "），请缩小选区。");
        }

        internal readonly struct Bounds
        {
            internal Bounds(int firstRow, int lastRow, int firstColumn, int lastColumn)
            {
                FirstRow = firstRow;
                LastRow = lastRow;
                FirstColumn = firstColumn;
                LastColumn = lastColumn;
                Validate();
            }

            internal int FirstRow { get; }
            internal int LastRow { get; }
            internal int FirstColumn { get; }
            internal int LastColumn { get; }
            internal int RowCount => LastRow - FirstRow + 1;
            internal int ColumnCount => LastColumn - FirstColumn + 1;
            internal long CellCount => (long)RowCount * ColumnCount;

            internal void Validate()
            {
                if (FirstRow <= 0 || FirstColumn <= 0 || LastRow < FirstRow || LastColumn < FirstColumn)
                {
                    throw new HostOperationException("选区或 UsedRange 的矩形边界无效。");
                }
            }
        }
    }
}
