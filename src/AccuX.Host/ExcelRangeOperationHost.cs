using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using AccuX.Core.Cells;
using AccuX.Core.Operations;
using Excel = Microsoft.Office.Interop.Excel;

namespace AccuX.Host
{
    /// <summary>
    /// Excel / WPS 表格宿主适配实现（规格 §5.2 / §8）。
    /// <para>
    /// 职责：宿主差异、COM 边界、批量 Value/Formula 读写、公式规范化、特殊区域识别、宿主状态管理。
    /// 所有 Excel/WPS 差异只允许出现在本工程；业务 Module 不接触 COM。
    /// </para>
    /// <para>
    /// COM 对象只在本类内部使用，不外泄到 Core / Module。
    /// V1 不对临时 RCW 显式调用 Marshal.ReleaseComObject（以实际兼容性验证为准），依赖 GC 回收。
    /// </para>
    /// </summary>
    public sealed partial class ExcelRangeOperationHost : IRangeOperationHost, IWorkbookDirectoryHost, ICellCommentHost, ICellMarkHost
    {
        private readonly Excel.Application _application;
        private readonly HostOptions _options;

        public ExcelRangeOperationHost(Excel.Application application, HostOptions options = null)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _options = options ?? new HostOptions();
            Context = HostDetector.Detect(application);
        }

        public HostOptions Options
        {
            get { return _options; }
        }

        public IHostContext Context { get; }

        /// <summary>
        /// 读取一次当前 Selection 并固化为 RangeTarget（规格 §19）。
        /// 一次 Command 只允许调用一次。
        /// </summary>
        public RangeTarget CaptureTarget()
        {
            var workbook = GetActiveWorkbook();
            var worksheet = GetActiveWorksheet(workbook);
            var selection = GetSelectionRange(worksheet);

            if (selection == null)
            {
                throw new HostOperationException("当前选区不是有效的单元格区域，请先选择一个数据区域。");
            }

            var areas = SafeAreaCount(selection);
            var rowCount = SafeRowCount(selection);
            var columnCount = SafeColumnCount(selection);
            var cellCount = (long)rowCount * columnCount;
            var address = SafeAddress(selection);

            if (rowCount <= 0 || columnCount <= 0)
            {
                throw new HostOperationException("当前选区为空，请先选择一个数据区域。");
            }

            if (areas > 1)
            {
                throw new HostOperationException("V1 不支持多区域选区，请选择单个连续区域后重试。");
            }

            if (cellCount > _options.MaxProcessCells)
            {
                throw new HostOperationException(
                    string.Format(
                        CultureInfo.CurrentCulture,
                        "选区包含 {0} 个单元格，超过 V1 上限 {1}。请缩小选区后重试。",
                        cellCount,
                        _options.MaxProcessCells));
            }

            var workbookKey = BuildWorkbookKey(workbook);
            var worksheetKey = SafeWorksheetName(worksheet);
            var worksheetName = SafeWorksheetName(worksheet);
            var containsMerged = DetectMergedCells(selection, rowCount, columnCount);

            return new RangeTarget(
                workbookKey,
                worksheetKey,
                worksheetName,
                address,
                rowCount,
                columnCount,
                cellCount,
                isMultiArea: false,
                containsMergedCells: containsMerged);
        }

        /// <summary>
        /// 写回前针对同一 RangeTarget 再次验证（规格 §19）。
        /// </summary>
        public WriteCheckResult ValidateWrite(RangeTarget target, RangeWritePlan writePlan)
        {
            if (target == null)
            {
                return WriteCheckResult.Failure("操作目标已丢失，请重新执行。");
            }

            try
            {
                var workbook = ResolveWorkbook(target);
                if (workbook == null)
                {
                    return WriteCheckResult.Failure("原工作簿已关闭，请重新执行。");
                }

                var worksheet = ResolveWorksheet(workbook, target);
                if (worksheet == null)
                {
                    return WriteCheckResult.Failure("原工作表已关闭或已重命名，请重新执行。");
                }

                Excel.Range range;
                try
                {
                    range = worksheet.Range[target.Address];
                }
                catch (Exception ex)
                {
                    return WriteCheckResult.Failure("原选区地址已失效：" + ex.Message);
                }

                if (SafeAreaCount(range) > 1)
                {
                    return WriteCheckResult.Failure("目标区域已变为多区域，无法安全写入。");
                }

                if (IsSheetProtected(worksheet))
                {
                    return WriteCheckResult.Failure("目标工作表处于保护状态，无法写入。请先取消保护。");
                }

                if (writePlan != null && !writePlan.IsEmpty)
                {
                    var specialAreaCheck = ValidateSpecialFormulaAreas(
                        range,
                        writePlan,
                        target.RowCount,
                        target.ColumnCount);
                    if (specialAreaCheck == null || !specialAreaCheck.CanWrite)
                    {
                        return specialAreaCheck ?? WriteCheckResult.Failure("无法确认目标区域的特殊公式状态。");
                    }
                }

                // 隐藏状态可能在读取后发生变化（例如被筛选器或其他宏改变）。
                // 如果待写单元格此时已落入隐藏行/列，终止写回，避免绕过“只处理可见单元格”的规则。
                if (writePlan != null && !writePlan.IsEmpty)
                {
                    var hiddenRows = ReadHiddenRows(worksheet, range, target.RowCount);
                    var hiddenColumns = ReadHiddenColumns(worksheet, range, target.ColumnCount);

                    foreach (var write in writePlan.Writes)
                    {
                        if (!IsInside(write.Row, write.Column, target.RowCount, target.ColumnCount))
                        {
                            continue;
                        }

                        if (hiddenRows[write.Row] || hiddenColumns[write.Column])
                        {
                            return WriteCheckResult.Failure("目标区域的隐藏行/列状态已发生变化，请重新执行。");
                        }
                    }
                }

                return WriteCheckResult.Success();
            }
            catch (Exception ex)
            {
                return WriteCheckResult.Failure("写入前校验失败：" + ex.Message);
            }
        }

        /// <summary>
        /// 创建宿主状态作用域。只保存并恢复本次真正修改过的 Application 状态。
        /// </summary>
        public IHostStateScope BeginStateScope(HostStateOptions options)
        {
            return new ExcelHostStateScope(_application, options ?? HostStateOptions.Default);
        }

        /// <summary>
        /// 读取 RangeTarget 的 NumberFormat 矩阵。
        /// </summary>
        public string[,] ReadNumberFormats(RangeTarget target)
        {
            var worksheet = ResolveWorksheetOrThrow(target);
            var range = worksheet.Range[target.Address];
            return ReadNumberFormatMatrix(range, target.RowCount, target.ColumnCount);
        }

        // ---------- 内部解析 ----------

        private Excel.Workbook GetActiveWorkbook()
        {
            var workbook = _application.ActiveWorkbook;
            if (workbook == null)
            {
                throw new HostOperationException("当前没有打开的工作簿。");
            }

            return workbook;
        }

        private Excel.Worksheet GetActiveWorksheet(Excel.Workbook workbook)
        {
            var sheet = workbook.ActiveSheet as Excel.Worksheet;
            if (sheet == null)
            {
                throw new HostOperationException("当前没有活动的工作表。");
            }

            return sheet;
        }

        private Excel.Range GetSelectionRange(Excel.Worksheet worksheet)
        {
            try
            {
                return _application.Selection as Excel.Range;
            }
            catch
            {
                return null;
            }
        }

        private Excel.Workbook ResolveWorkbook(RangeTarget target)
        {
            if (target == null || string.IsNullOrEmpty(target.WorkbookKey))
            {
                return null;
            }

            foreach (Excel.Workbook workbook in _application.Workbooks)
            {
                if (string.Equals(BuildWorkbookKey(workbook), target.WorkbookKey, StringComparison.OrdinalIgnoreCase))
                {
                    return workbook;
                }
            }

            return null;
        }

        private Excel.Worksheet ResolveWorksheet(Excel.Workbook workbook, RangeTarget target)
        {
            if (workbook == null || target == null || string.IsNullOrEmpty(target.WorksheetKey))
            {
                return null;
            }

            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (string.Equals(SafeWorksheetName(sheet), target.WorksheetKey, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }
            }

            return null;
        }

        private Excel.Worksheet ResolveWorksheetOrThrow(RangeTarget target)
        {
            var workbook = ResolveWorkbook(target);
            if (workbook == null)
            {
                throw new HostOperationException("原工作簿已关闭，请重新执行。");
            }

            var worksheet = ResolveWorksheet(workbook, target);
            if (worksheet == null)
            {
                throw new HostOperationException("原工作表已关闭或已重命名，请重新执行。");
            }

            return worksheet;
        }

        private static string BuildWorkbookKey(Excel.Workbook workbook)
        {
            try
            {
                var fullName = workbook.FullName;
                if (!string.IsNullOrWhiteSpace(fullName))
                {
                    return fullName;
                }
            }
            catch
            {
                // 某些宿主在特定状态下可能取不到 FullName，回退到名称。
            }

            try
            {
                return workbook.Name ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeWorksheetName(Excel.Worksheet worksheet)
        {
            try
            {
                return worksheet.Name ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeAddress(Excel.Range range)
        {
            try
            {
                return range.Address;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static int SafeRowCount(Excel.Range range)
        {
            try
            {
                return Convert.ToInt32(range.Rows.Count, CultureInfo.InvariantCulture);
            }
            catch
            {
                return 0;
            }
        }

        private static int SafeColumnCount(Excel.Range range)
        {
            try
            {
                return Convert.ToInt32(range.Columns.Count, CultureInfo.InvariantCulture);
            }
            catch
            {
                return 0;
            }
        }

        private static int SafeAreaCount(Excel.Range range)
        {
            try
            {
                return Convert.ToInt32(range.Areas.Count, CultureInfo.InvariantCulture);
            }
            catch
            {
                return 1;
            }
        }

        private static bool IsSheetProtected(Excel.Worksheet worksheet)
        {
            try
            {
                return worksheet.ProtectContents;
            }
            catch
            {
                return false;
            }
        }

        private static bool DetectMergedCells(Excel.Range range, int rowCount, int columnCount)
        {
            try
            {
                var merged = range.MergeCells;
                if (merged is bool boolValue)
                {
                    return boolValue;
                }

                // 混合状态返回 null/DBNull：保守判定为包含合并单元格。
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
