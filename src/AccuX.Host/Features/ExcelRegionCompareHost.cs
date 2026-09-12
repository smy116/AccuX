using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using AccuX.Core.Operations;
using Excel = Microsoft.Office.Interop.Excel;

namespace AccuX.Host.Features
{
    /// <summary>
    /// 区域对比的 Excel 宿主实现。所有 COM 解析、着色和导出都封装在此处。
    /// </summary>
    public sealed class ExcelRegionCompareHost : ExcelHostBase, IRegionCompareHost
    {
        private readonly ExcelRangeOperationHost _rangeHost;
        private readonly HostOptions _options;
        private readonly Dictionary<string, Excel.Workbook> _workbookSessions = new Dictionary<string, Excel.Workbook>(StringComparer.Ordinal);

        public ExcelRegionCompareHost(Excel.Application application, HostOptions options)
            : base(application)
        {
            _options = options ?? new HostOptions();
            if (_options.MaxProcessCells <= 0)
            {
                _options.MaxProcessCells = 500000;
            }
            if (_options.LargeSelectionWarning <= 0)
            {
                _options.LargeSelectionWarning = 100000;
            }

            _rangeHost = new ExcelRangeOperationHost(application, _options, ResolveSessionWorkbook);
        }

        public ExcelRegionCompareHost(Excel.Application application, long maxProcessCells = 500000)
            : this(application, new HostOptions { MaxProcessCells = maxProcessCells })
        {
        }

        public HostOptions Options { get { return _options; } }

        protected override Excel.Workbook ResolveWorkbook(RangeTarget target)
        {
            return ResolveSessionWorkbook(target);
        }

        public RangeTarget CaptureCurrentTarget()
        {
            var workbook = GetActiveWorkbook();
            var worksheet = GetActiveWorksheet(workbook);
            var selection = GetSelectionRange(worksheet);
            if (selection == null)
            {
                throw new HostOperationException("当前选区不是有效的单元格区域，请先选择一个连续区域。");
            }

            var areas = SafeAreaCount(selection);
            var rows = SafeRowCount(selection);
            var columns = SafeColumnCount(selection);
            var cells = (long)rows * columns;
            if (areas > 1) throw new HostOperationException("区域对比不支持多区域选区，请选择单个连续区域。");
            if (rows <= 0 || columns <= 0) throw new HostOperationException("当前选区为空。");
            if (cells > _options.MaxProcessCells)
            {
                throw new HostOperationException("当前选区包含 " + cells + " 个单元格，超过区域对比上限 " + _options.MaxProcessCells + "。");
            }
            if (HasMergedCells(selection))
            {
                throw new HostOperationException("区域对比不支持包含合并单元格的选区。");
            }

            var identityToken = Guid.NewGuid().ToString("N");
            _workbookSessions[identityToken] = workbook;
            return new RangeTarget(
                BuildWorkbookKey(workbook),
                BuildWorksheetKey(worksheet),
                SafeWorksheetName(worksheet),
                SafeAddress(selection),
                rows,
                columns,
                cells,
                false,
                false,
                identityToken);
        }

        private Excel.Workbook ResolveSessionWorkbook(RangeTarget target)
        {
            if (target == null || string.IsNullOrEmpty(target.IdentityToken))
            {
                return base.ResolveWorkbook(target);
            }

            if (!_workbookSessions.TryGetValue(target.IdentityToken, out var captured))
            {
                return null;
            }

            try
            {
                foreach (Excel.Workbook open in Application.Workbooks)
                {
                    if (ReferenceEquals(open, captured) || SameComObject(open, captured))
                    {
                        return open;
                    }
                }
            }
            catch
            {
                return null;
            }

            return null;
        }

        private static bool SameComObject(object first, object second)
        {
            IntPtr firstUnknown = IntPtr.Zero;
            IntPtr secondUnknown = IntPtr.Zero;
            try
            {
                firstUnknown = Marshal.GetIUnknownForObject(first);
                secondUnknown = Marshal.GetIUnknownForObject(second);
                return firstUnknown != IntPtr.Zero && firstUnknown == secondUnknown;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (firstUnknown != IntPtr.Zero) Marshal.Release(firstUnknown);
                if (secondUnknown != IntPtr.Zero) Marshal.Release(secondUnknown);
            }
        }

        public RangeReadResult ReadRegion(RangeTarget target)
        {
            return _rangeHost.Read(target);
        }

        public bool ValidateTarget(RangeTarget target, bool requireWritable, out string message)
        {
            message = null;
            if (target == null)
            {
                message = "区域目标为空，请重新捕获选区。";
                return false;
            }

            try
            {
                var workbook = ResolveWorkbook(target);
                if (workbook == null) { message = "原工作簿已关闭，请重新捕获选区。"; return false; }
                var worksheet = ResolveWorksheet(workbook, target);
                if (worksheet == null) { message = "原工作表已关闭或已重命名，请重新捕获选区。"; return false; }
                var range = worksheet.Range[target.Address];
                if (range == null || SafeAreaCount(range) != 1) { message = "原区域已失效，请重新捕获选区。"; return false; }
                if (SafeRowCount(range) != target.RowCount || SafeColumnCount(range) != target.ColumnCount)
                {
                    message = "原区域大小已发生变化，请重新捕获选区。";
                    return false;
                }
                if (HasMergedCells(range)) { message = "原区域已包含合并单元格，无法继续。"; return false; }
                if (requireWritable && IsSheetProtected(worksheet))
                {
                    message = "目标工作表处于保护状态，无法修改底色。请先取消保护。";
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                message = "验证区域失败：" + ex.Message;
                return false;
            }
        }

        public bool ValidatePositions(
            RangeTarget target,
            IReadOnlyList<RegionCompareCellPosition> positions,
            out string message)
        {
            message = null;
            if (positions == null || positions.Count == 0) return true;
            if (!ValidateTarget(target, true, out message)) return false;
            try
            {
                var worksheet = ResolveWorksheetOrThrow(target);
                var range = worksheet.Range[target.Address];
                var hiddenRows = ReadHiddenRows(worksheet, range, target.RowCount);
                var hiddenColumns = ReadHiddenColumns(worksheet, range, target.ColumnCount);
                GroupPositions(positions, target.RowCount, target.ColumnCount, hiddenRows, hiddenColumns);
                return true;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return false;
            }
        }

        public long ApplyBackgroundColor(
            RangeTarget target,
            IReadOnlyList<RegionCompareCellPosition> positions,
            string hexColor)
        {
            return ApplyColor(target, positions, ExcelComHelper.ParseOleColor(hexColor), false);
        }

        public long ClearBackgroundColor(
            RangeTarget target,
            IReadOnlyList<RegionCompareCellPosition> positions)
        {
            return ApplyColor(target, positions, 0, true);
        }

        public void ExportResults(RegionCompareExportData data)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            var workbook = Application.Workbooks.Add();
            try
            {
                using (CreateStateScope(new HostStateOptions
                {
                    DisableScreenUpdating = true,
                    DisableEvents = true,
                    DisableDisplayAlerts = true
                }))
                {
                    // Excel 的默认工作表数量可由用户设置，先收敛到一张再补齐四张结果表。
                    while (workbook.Worksheets.Count > 1)
                    {
                        ((Excel.Worksheet)workbook.Worksheets[2]).Delete();
                    }

                    var first = (Excel.Worksheet)workbook.Worksheets[1];
                    first.Name = "区域1独有";
                    var second = (Excel.Worksheet)workbook.Worksheets.Add(Type.Missing, first, Type.Missing, Type.Missing);
                    second.Name = "区域2独有";
                    var same = (Excel.Worksheet)workbook.Worksheets.Add(Type.Missing, second, Type.Missing, Type.Missing);
                    same.Name = "相同项";
                    var detail = (Excel.Worksheet)workbook.Worksheets.Add(Type.Missing, same, Type.Missing, Type.Missing);
                    detail.Name = "来源明细";

                    WriteCategory(first, data, "区域1独有");
                    WriteCategory(second, data, "区域2独有");
                    WriteCategory(same, data, "相同项");
                    WriteDetails(detail, data);
                    workbook.Activate();
                }
            }
            catch
            {
                // 保留新工作簿，便于用户在出现单表写入问题时查看已生成内容。
                throw;
            }
        }

        public void ReleaseSession()
        {
            _workbookSessions.Clear();
        }

        private long ApplyColor(
            RangeTarget target,
            IReadOnlyList<RegionCompareCellPosition> positions,
            int oleColor,
            bool clear)
        {
            if (positions == null || positions.Count == 0) return 0;
            if (!ValidateTarget(target, true, out var message)) throw new HostOperationException(message);
            var worksheet = ResolveWorksheetOrThrow(target);
            var range = worksheet.Range[target.Address];
            var hiddenRows = ReadHiddenRows(worksheet, range, target.RowCount);
            var hiddenColumns = ReadHiddenColumns(worksheet, range, target.ColumnCount);
            var grouped = GroupPositions(positions, target.RowCount, target.ColumnCount, hiddenRows, hiddenColumns);

            // 先完成所有位置、可见性和范围检查，再执行任何写入。
            using (CreateStateScope(new HostStateOptions
            {
                DisableScreenUpdating = true,
                DisableEvents = true,
                DisableDisplayAlerts = true
            }))
            {
                foreach (var group in grouped)
                {
                    Excel.Range block = ((Excel.Range)range.Cells[group.Row + 1, group.Column + 1]).Resize[1, group.Length];
                    if (clear)
                    {
                        block.Interior.Pattern = Excel.XlPattern.xlPatternNone;
                    }
                    else
                    {
                        block.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                        block.Interior.Color = oleColor;
                    }
                }
            }

            return grouped.Count == 0 ? 0 : CountGroups(grouped);
        }

        private static List<PositionRun> GroupPositions(
            IReadOnlyList<RegionCompareCellPosition> positions,
            int rows,
            int columns,
            bool[] hiddenRows,
            bool[] hiddenColumns)
        {
            var byRow = new Dictionary<int, List<int>>();
            foreach (var position in positions)
            {
                if (position == null || position.Row < 0 || position.Row >= rows || position.Column < 0 || position.Column >= columns)
                {
                    throw new HostOperationException("对比结果包含超出原区域范围的位置，请重新对比。");
                }
                if (hiddenRows[position.Row] || hiddenColumns[position.Column])
                {
                    throw new HostOperationException("目标区域的隐藏行/列状态已发生变化，请重新对比。");
                }
                if (!byRow.TryGetValue(position.Row, out var list)) { list = new List<int>(); byRow.Add(position.Row, list); }
                if (!list.Contains(position.Column)) list.Add(position.Column);
            }

            var result = new List<PositionRun>();
            foreach (var pair in byRow)
            {
                pair.Value.Sort();
                var start = pair.Value[0];
                var last = start;
                for (var i = 1; i < pair.Value.Count; i++)
                {
                    if (pair.Value[i] == last + 1) { last++; continue; }
                    result.Add(new PositionRun(pair.Key, start, last - start + 1));
                    start = last = pair.Value[i];
                }
                result.Add(new PositionRun(pair.Key, start, last - start + 1));
            }
            return result;
        }

        private static long CountGroups(List<PositionRun> groups)
        {
            long count = 0;
            foreach (var group in groups) count += group.Length;
            return count;
        }

        private static bool HasMergedCells(Excel.Range range)
        {
            try
            {
                var merged = range.MergeCells;
                return merged is bool value ? value : true;
            }
            catch { return true; }
        }

        private static void WriteCategory(Excel.Worksheet sheet, RegionCompareExportData data, string category)
        {
            var rows = new List<RegionCompareExportRow>();
            foreach (var row in data.Rows) if (row != null && row.Category == category) rows.Add(row);
            WriteRows(sheet, rows, data.Summary);
        }

        private static void WriteDetails(Excel.Worksheet sheet, RegionCompareExportData data)
        {
            WriteRows(sheet, new List<RegionCompareExportRow>(data.Rows), data.Summary);
        }

        private static void WriteRows(Excel.Worksheet sheet, List<RegionCompareExportRow> rows, string summary)
        {
            var values = new object[rows.Count + 2, 13];
            values[0, 0] = "分类"; values[0, 1] = "值"; values[0, 2] = "类型"; values[0, 3] = "区域1次数";
            values[0, 4] = "区域2次数"; values[0, 5] = "区域1工作簿"; values[0, 6] = "区域1工作表";
            values[0, 7] = "区域1范围"; values[0, 8] = "区域1位置"; values[0, 9] = "区域2工作簿";
            values[0, 10] = "区域2工作表"; values[0, 11] = "区域2范围"; values[0, 12] = "区域2位置";
            for (var i = 0; i < rows.Count; i++)
            {
                var row = rows[i];
                values[i + 1, 0] = SafeText(row.Category);
                values[i + 1, 1] = SafeText(row.Value);
                values[i + 1, 2] = SafeText(row.CellType);
                values[i + 1, 3] = row.FirstCount;
                values[i + 1, 4] = row.SecondCount;
                values[i + 1, 5] = SafeText(row.FirstWorkbook);
                values[i + 1, 6] = SafeText(row.FirstWorksheet);
                values[i + 1, 7] = SafeText(row.FirstRange);
                values[i + 1, 8] = SafeText(row.FirstLocations);
                values[i + 1, 9] = SafeText(row.SecondWorkbook);
                values[i + 1, 10] = SafeText(row.SecondWorksheet);
                values[i + 1, 11] = SafeText(row.SecondRange);
                values[i + 1, 12] = SafeText(row.SecondLocations);
            }
            values[rows.Count + 1, 0] = SafeText(summary);
            var end = (Excel.Range)sheet.Cells[rows.Count + 2, 13];
            var start = (Excel.Range)sheet.Cells[1, 1];
            sheet.Range[start, end].Value2 = values;
            sheet.Columns.AutoFit();
        }

        private static string SafeText(string value)
        {
            if (string.IsNullOrEmpty(value)) return value ?? string.Empty;
            return value.StartsWith("=", StringComparison.Ordinal) ? "'" + value : value;
        }

        private struct PositionRun
        {
            public PositionRun(int row, int column, int length) { Row = row; Column = column; Length = length; }
            public int Row { get; }
            public int Column { get; }
            public int Length { get; }
        }
    }
}
