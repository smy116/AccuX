using System;
using System.Collections.Generic;
using AccuX.Core.Operations;
using Excel = Microsoft.Office.Interop.Excel;

namespace AccuX.Host.Features
{
    /// <summary>
    /// 可见单元格底色标记功能的宿主实现：只负责 COM 写入与可见性处理，颜色由业务模块提供。
    /// </summary>
    public sealed class ExcelCellMarkHost : ExcelHostBase, ICellMarkHost
    {
        public ExcelCellMarkHost(Excel.Application application)
            : base(application)
        {
        }

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

            var oleColor = ExcelComHelper.ParseOleColor(hexColor);
            if (target.Areas.Count == 0 || (target.IsMultiArea && target.Areas.Count == 1))
            {
                throw new HostOperationException("选区缺少有效的区域信息，请重新选择。");
            }

            var worksheet = ResolveWorksheetOrThrow(target);
            if (IsSheetProtected(worksheet))
            {
                throw new HostOperationException("目标工作表处于保护状态，无法标记底色。请先取消保护。");
            }

            long visibleCellCount = 0;
            var visibleRanges = new List<Excel.Range>();
            // 所有区域完成解析和校验之后才能开始着色。
            foreach (var area in target.Areas)
            {
                Excel.Range range;
                try
                {
                    range = worksheet.Range[area.Address];
                    if (range == null || range.Areas.Count != 1
                        || range.Rows.Count != area.RowCount || range.Columns.Count != area.ColumnCount)
                    {
                        throw new HostOperationException("原选区地址或大小已发生变化，请重新选择后重试。");
                    }

                    if (!(range.MergeCells is bool merged) || merged)
                    {
                        throw new HostOperationException("选区包含合并单元格，无法标记底色。");
                    }
                }
                catch (HostOperationException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    throw new HostOperationException("无法解析原选区地址：" + ex.Message, ex);
                }

                var hiddenRows = ReadHiddenRows(worksheet, range, area.RowCount);
                var hiddenColumns = ReadHiddenColumns(worksheet, range, area.ColumnCount);
                visibleCellCount += CountVisibleCells(hiddenRows, hiddenColumns);
                visibleRanges.AddRange(BuildVisibleRanges(range, hiddenRows, hiddenColumns));
            }

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
