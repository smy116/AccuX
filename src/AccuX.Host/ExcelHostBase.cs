using System;
using System.Globalization;
using AccuX.Core.Operations;
using Excel = Microsoft.Office.Interop.Excel;

namespace AccuX.Host
{
    /// <summary>
    /// 各宿主能力实现类共享的 COM 基础设施（规格 §5.2 / §8）。
    /// <para>
    /// 只提供机制：Application 引用、工作簿与工作表解析、安全的属性读取、宿主状态作用域创建。
    /// 不包含任何业务文案或外观决定；COM 对象只在 AccuX.Host 内部使用，不外泄到 Core / Module。
    /// </para>
    /// <para>
    /// 能力实现类保持“一个 Core 窄接口一个类”：范围读写、目录、批注、标记各自独立，
    /// 需要 COM 机制时从本类继承，而不是继续堆到同一个类里。
    /// </para>
    /// </summary>
    public abstract class ExcelHostBase
    {
        protected ExcelHostBase(Excel.Application application)
        {
            Application = application ?? throw new ArgumentNullException(nameof(application));
        }

        protected Excel.Application Application { get; }

        /// <summary>
        /// 创建宿主状态作用域；只保存并恢复本次真正修改过的 Application 状态。
        /// </summary>
        protected IHostStateScope CreateStateScope(HostStateOptions options)
        {
            return new ExcelHostStateScope(Application, options ?? HostStateOptions.Default);
        }

        protected Excel.Workbook GetActiveWorkbook()
        {
            var workbook = Application.ActiveWorkbook;
            if (workbook == null)
            {
                throw new HostOperationException("当前没有打开的工作簿。");
            }

            return workbook;
        }

        protected Excel.Worksheet GetActiveWorksheet(Excel.Workbook workbook)
        {
            var sheet = workbook.ActiveSheet as Excel.Worksheet;
            if (sheet == null)
            {
                throw new HostOperationException("当前没有活动的工作表。");
            }

            return sheet;
        }

        protected Excel.Range GetSelectionRange(Excel.Worksheet worksheet)
        {
            try
            {
                return Application.Selection as Excel.Range;
            }
            catch
            {
                return null;
            }
        }

        protected Excel.Workbook ResolveWorkbook(RangeTarget target)
        {
            if (target == null || string.IsNullOrEmpty(target.WorkbookKey))
            {
                return null;
            }

            foreach (Excel.Workbook workbook in Application.Workbooks)
            {
                if (string.Equals(BuildWorkbookKey(workbook), target.WorkbookKey, StringComparison.OrdinalIgnoreCase))
                {
                    return workbook;
                }
            }

            return null;
        }

        protected Excel.Worksheet ResolveWorksheet(Excel.Workbook workbook, RangeTarget target)
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

        protected Excel.Worksheet ResolveWorksheetOrThrow(RangeTarget target)
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

        protected static string BuildWorkbookKey(Excel.Workbook workbook)
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

        protected static string SafeWorksheetName(Excel.Worksheet worksheet)
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

        protected static string SafeAddress(Excel.Range range)
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

        protected static int SafeRowCount(Excel.Range range)
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

        protected static int SafeColumnCount(Excel.Range range)
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

        protected static int SafeAreaCount(Excel.Range range)
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

        protected static bool IsSheetProtected(Excel.Worksheet worksheet)
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

        /// <summary>
        /// 读取目标选区覆盖的行隐藏状态。
        /// 优先读取选区覆盖的完整行集合；只有明确不是“全部隐藏”时才逐行回退，
        /// 以避免兼容宿主把混合状态误报为“全部可见”。
        /// </summary>
        protected static bool[] ReadHiddenRows(Excel.Worksheet worksheet, Excel.Range range, int rowCount)
        {
            var result = new bool[rowCount];
            object combinedHidden = null;

            try
            {
                // Hidden 要求 Range 覆盖完整行。range.Rows 仍然只覆盖选区的列片段，
                // 在 Excel/WPS 中对混合可见性可能返回 Null、抛异常，或被兼容宿主折叠为 False。
                combinedHidden = range.EntireRow.Hidden;
            }
            catch
            {
                // 回退到逐行读取。
            }

            // 兼容宿主可能把“部分隐藏、部分可见”折叠成 False；False 不能证明所有行都可见，
            // 因此只有明确的 True 才走批量快路径，其他情况统一逐行读取。
            if (TryConvertBooleanValue(combinedHidden, out var allHidden) && allHidden)
            {
                Fill(result, true);
                return result;
            }

            var firstRow = range.Row;
            for (var row = 0; row < rowCount; row++)
            {
                var absoluteRow = firstRow + row;
                try
                {
                    var rowRange = (Excel.Range)worksheet.Rows.get_Item(absoluteRow, Type.Missing);
                    result[row] = ReadHiddenValue(rowRange, "行", absoluteRow);
                }
                catch (HostOperationException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    throw new HostOperationException(
                        "无法读取第 " + absoluteRow + " 行的隐藏状态。",
                        ex);
                }
            }

            return result;
        }

        /// <summary>
        /// 读取目标选区覆盖的列隐藏状态。
        /// 优先读取选区覆盖的完整列集合；只有明确不是“全部隐藏”时才逐列回退，
        /// 以避免兼容宿主把混合状态误报为“全部可见”。
        /// </summary>
        protected static bool[] ReadHiddenColumns(Excel.Worksheet worksheet, Excel.Range range, int columnCount)
        {
            var result = new bool[columnCount];
            object combinedHidden = null;

            try
            {
                // Hidden 要求 Range 覆盖完整列。range.Columns 仍然只覆盖选区的行片段，
                // 在 Excel/WPS 中对混合可见性可能返回 Null、抛异常，或被兼容宿主折叠为 False。
                combinedHidden = range.EntireColumn.Hidden;
            }
            catch
            {
                // 回退到逐列读取。
            }

            // 兼容宿主可能把“部分隐藏、部分可见”折叠成 False；False 不能证明所有列都可见，
            // 因此只有明确的 True 才走批量快路径，其他情况统一逐列读取。
            if (TryConvertBooleanValue(combinedHidden, out var allHidden) && allHidden)
            {
                Fill(result, true);
                return result;
            }

            var firstColumn = range.Column;
            for (var column = 0; column < columnCount; column++)
            {
                var absoluteColumn = firstColumn + column;
                try
                {
                    var columnRange = (Excel.Range)worksheet.Columns.get_Item(absoluteColumn, Type.Missing);
                    result[column] = ReadHiddenValue(columnRange, "列", absoluteColumn);
                }
                catch (HostOperationException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    throw new HostOperationException(
                        "无法读取第 " + absoluteColumn + " 列的隐藏状态。",
                        ex);
                }
            }

            return result;
        }

        private static bool ReadHiddenValue(Excel.Range dimensionRange, string dimensionName, int index)
        {
            if (dimensionRange == null)
            {
                throw new HostOperationException(
                    "无法确定第 " + index + " " + dimensionName + "的隐藏状态。");
            }

            object raw = dimensionRange.Hidden;
            if (TryConvertBooleanValue(raw, out var hidden))
            {
                return hidden;
            }

            throw new HostOperationException(
                "无法确定第 " + index + " " + dimensionName + "的隐藏状态。");
        }

        protected static bool TryConvertBooleanValue(object raw, out bool hidden)
        {
            if (raw is bool boolValue)
            {
                hidden = boolValue;
                return true;
            }

            // 某些 COM 兼容宿主可能将 VARIANT_BOOL 映射为整数。
            switch (raw)
            {
                case sbyte sbyteValue:
                    hidden = sbyteValue != 0;
                    return true;
                case byte byteValue:
                    hidden = byteValue != 0;
                    return true;
                case short shortValue:
                    hidden = shortValue != 0;
                    return true;
                case ushort ushortValue:
                    hidden = ushortValue != 0;
                    return true;
                case int intValue:
                    hidden = intValue != 0;
                    return true;
                case uint uintValue:
                    hidden = uintValue != 0;
                    return true;
                case long longValue:
                    hidden = longValue != 0;
                    return true;
                case ulong ulongValue:
                    hidden = ulongValue != 0;
                    return true;
                default:
                    hidden = false;
                    return false;
            }
        }

        private static void Fill(bool[] values, bool value)
        {
            for (var i = 0; i < values.Length; i++)
            {
                values[i] = value;
            }
        }
    }
}
