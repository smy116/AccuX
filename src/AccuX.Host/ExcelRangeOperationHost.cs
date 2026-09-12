using System;
using System.Globalization;
using AccuX.Core.Operations;
using Excel = Microsoft.Office.Interop.Excel;

namespace AccuX.Host
{
    /// <summary>
    /// 范围操作宿主实现（规格 §5.2 / §8）：批量 Value/Formula 读写、公式规范化、特殊区域识别、宿主状态管理。
    /// <para>
    /// 只实现 <see cref="IRangeOperationHost"/>；目录、批注、标记等具体功能的 COM 实现位于
    /// Features 目录下各自的宿主类，共享机制来自 <see cref="ExcelHostBase"/>。
    /// </para>
    /// <para>
    /// COM 对象只在本工程内部使用，不外泄到 Core / Module。
    /// V1 不对临时 RCW 显式调用 Marshal.ReleaseComObject（以实际兼容性验证为准），依赖 GC 回收。
    /// </para>
    /// </summary>
    public sealed partial class ExcelRangeOperationHost : ExcelHostBase, IRangeOperationHost
    {
        private readonly HostOptions _options;
        private readonly Func<RangeTarget, Excel.Workbook> _workbookResolver;

        public ExcelRangeOperationHost(
            Excel.Application application,
            HostOptions options = null,
            Func<RangeTarget, Excel.Workbook> workbookResolver = null)
            : base(application)
        {
            _options = options ?? new HostOptions();
            _workbookResolver = workbookResolver;
            Context = HostDetector.Detect(application);
        }

        protected override Excel.Workbook ResolveWorkbook(RangeTarget target)
        {
            return _workbookResolver == null ? base.ResolveWorkbook(target) : _workbookResolver(target);
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
            var originalRowCount = SafeRowCount(selection);
            var originalColumnCount = SafeColumnCount(selection);

            if (originalRowCount <= 0 || originalColumnCount <= 0)
            {
                throw new HostOperationException("当前选区为空，请先选择一个数据区域。");
            }

            if (areas > 1)
            {
                throw new HostOperationException("V1 不支持多区域选区，请选择单个连续区域后重试。");
            }

            // 批量功能只在当前 Selection 与工作表原生 UsedRange 的交集内生效。
            // Selection 本身仍由 GetSelectionRange 原样获取，批注助手等单元格功能不受影响。
            selection = RestrictSelectionToUsedRange(worksheet, selection);
            var rowCount = SafeRowCount(selection);
            var columnCount = SafeColumnCount(selection);
            var cellCount = (long)rowCount * columnCount;
            var address = SafeAddress(selection);

            if (rowCount <= 0 || columnCount <= 0)
            {
                throw new HostOperationException("UsedRange 内没有可处理的单元格，请重新选择数据区域。");
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
            var worksheetKey = BuildWorksheetKey(worksheet);
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
            return CreateStateScope(options);
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
