using System;
using System.Collections.Generic;
using System.Globalization;
using AccuX.Core.Operations;
using Excel = Microsoft.Office.Interop.Excel;

namespace AccuX.Host
{
    public sealed partial class ExcelRangeOperationHost
    {
        /// <summary>
        /// 将颜色应用到目标选区中的可见单元格。
        /// 通过行列连续区段批量写入，避免逐 Cell COM 调用。
        /// </summary>
        public long ApplyVisibleBackgroundColor(RangeTarget target, string hexColor)
        {
            if (target == null)
            {
                throw new HostOperationException("标记目标已丢失，请重新执行。");
            }

            var oleColor = ParseOleColor(hexColor);
            if (target.IsMultiArea || target.RowCount <= 0 || target.ColumnCount <= 0)
            {
                throw new HostOperationException("当前选区不是有效的连续单元格区域，请重新选择后重试。");
            }

            var worksheet = ResolveWorksheetOrThrow(target);
            Excel.Range range;
            try
            {
                range = worksheet.Range[target.Address];
            }
            catch (Exception ex)
            {
                throw new HostOperationException("无法解析原选区地址：" + ex.Message, ex);
            }

            if (range == null || SafeAreaCount(range) > 1)
            {
                throw new HostOperationException("原选区地址已失效，请重新执行。");
            }

            if (SafeRowCount(range) != target.RowCount || SafeColumnCount(range) != target.ColumnCount)
            {
                throw new HostOperationException("原选区大小已发生变化，请重新选择后重试。");
            }

            if (IsSheetProtected(worksheet))
            {
                throw new HostOperationException("目标工作表处于保护状态，无法标记底色。请先取消保护。");
            }

            // 与 RangeOperationHost.Read 共用经过兼容性处理的行列隐藏状态读取逻辑。
            var hiddenRows = ReadHiddenRows(worksheet, range, target.RowCount);
            var hiddenColumns = ReadHiddenColumns(worksheet, range, target.ColumnCount);
            var visibleCellCount = CountVisibleCells(hiddenRows, hiddenColumns);
            if (visibleCellCount == 0)
            {
                return 0;
            }
            
            // 先解析出所有待写的连续区域，再开始修改，尽量遵守 fail-before-write。
            var visibleRanges = BuildVisibleRanges(range, hiddenRows, hiddenColumns);
            try
            {
                foreach (var visibleRange in visibleRanges)
                {
                    ApplyBackground(visibleRange, oleColor);
                }
            }
            catch (Exception ex)
            {
                throw new HostOperationException("标记选区底色失败：" + ex.Message, ex);
            }

            return visibleCellCount;
        }

        internal static int ParseOleColor(string hexColor)
        {
            if (string.IsNullOrWhiteSpace(hexColor)
                || hexColor.Length != 7
                || hexColor[0] != '#')
            {
                throw new HostOperationException("标记颜色无效，必须使用 #RRGGBB 格式。");
            }

            try
            {
                var red = int.Parse(hexColor.Substring(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                var green = int.Parse(hexColor.Substring(3, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                var blue = int.Parse(hexColor.Substring(5, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                return red + (green << 8) + (blue << 16);
            }
            catch (FormatException ex)
            {
                throw new HostOperationException("标记颜色无效，必须使用 #RRGGBB 格式。", ex);
            }
            catch (OverflowException ex)
            {
                throw new HostOperationException("标记颜色无效，必须使用 #RRGGBB 格式。", ex);
            }
        }

        private static long CountVisibleCells(bool[] hiddenRows, bool[] hiddenColumns)
        {
            long count = 0;
            for (var row = 0; row < hiddenRows.Length; row++)
            {
                if (hiddenRows[row])
                {
                    continue;
                }

                for (var column = 0; column < hiddenColumns.Length; column++)
                {
                    if (!hiddenColumns[column])
                    {
                        count++;
                    }
                }
            }

            return count;
        }

        private static List<Excel.Range> BuildVisibleRanges(
            Excel.Range range,
            bool[] hiddenRows,
            bool[] hiddenColumns)
        {
            var result = new List<Excel.Range>();
            var visibleRowRuns = BuildVisibleRuns(hiddenRows);
            var visibleColumnRuns = BuildVisibleRuns(hiddenColumns);

            foreach (var rowRun in visibleRowRuns)
            {
                foreach (var columnRun in visibleColumnRuns)
                {
                    try
                    {
                        var visibleRange = ((Excel.Range)range.Cells[rowRun.Start + 1, columnRun.Start + 1])
                            .Resize[rowRun.Length, columnRun.Length];
                        if (visibleRange == null)
                        {
                            throw new HostOperationException("无法确定可见标记区域，请重新执行。");
                        }

                        result.Add(visibleRange);
                    }
                    catch (HostOperationException)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        throw new HostOperationException("无法确定可见标记区域：" + ex.Message, ex);
                    }
                }
            }

            return result;
        }

        private static List<VisibleRun> BuildVisibleRuns(bool[] hidden)
        {
            var result = new List<VisibleRun>();
            var index = 0;
            while (index < hidden.Length)
            {
                if (hidden[index])
                {
                    index++;
                    continue;
                }

                var start = index;
                while (index < hidden.Length && !hidden[index])
                {
                    index++;
                }

                result.Add(new VisibleRun(start, index - start));
            }

            return result;
        }

        private static void ApplyBackground(Excel.Range range, int oleColor)
        {
            range.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            range.Interior.Color = oleColor;
        }

        private struct VisibleRun
        {
            public VisibleRun(int start, int length)
            {
                Start = start;
                Length = length;
            }

            public int Start { get; }

            public int Length { get; }
        }
    }
}
