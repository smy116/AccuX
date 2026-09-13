using System;
using System.Collections.Generic;
using System.IO;
using AccuX.Core.Operations;
using Excel = Microsoft.Office.Interop.Excel;

namespace AccuX.Host.Features
{
    /// <summary>
    /// “生成目录”功能的宿主实现。
    /// <para>
    /// 只负责 COM 机制：查找 / 删除 / 新建工作表、批量写入、超链接与样式落地、失败回滚。
    /// 目录的文案与外观（标题、表头、配色、列宽、行高、字号）全部来自业务模块提供的
    /// <see cref="DirectoryOptions"/>，本类不内置任何目录专属常量。
    /// </para>
    /// </summary>
    public sealed class ExcelWorkbookDirectoryHost : ExcelHostBase, IWorkbookDirectoryHost
    {
        private const int ColumnCount = 3;

        public ExcelWorkbookDirectoryHost(Excel.Application application)
            : base(application)
        {
        }

        /// <summary>
        /// 判断当前活动工作簿中是否已存在指定名称的工作表。
        /// </summary>
        public bool DirectoryWorksheetExists(string worksheetName)
        {
            var workbook = GetActiveWorkbook();
            return FindWorksheet(workbook, worksheetName) != null;
        }

        /// <summary>
        /// 在当前工作簿最前面生成目录工作表；replaceExisting 为 true 时先完整生成新表，
        /// 再替换同名旧表。生成过程失败时尽量恢复原目录；若旧表删除本身失败，
        /// 则保留新表与旧表备份，避免破坏性清理造成数据丢失。
        /// </summary>
        public int GenerateDirectory(DirectoryOptions options, bool replaceExisting)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            ValidateOptions(options);

            Excel.Workbook workbook = null;
            Excel.Worksheet existingWorksheet = null;
            Excel.Worksheet directoryWorksheet = null;
            var originalWorksheetName = string.Empty;
            var existingRenamed = false;
            var existingDeleteAttempted = false;
            var completed = false;

            try
            {
                workbook = GetActiveWorkbook();
                existingWorksheet = FindWorksheet(workbook, options.WorksheetName);
                if (existingWorksheet != null && !replaceExisting)
                {
                    throw new HostOperationException(
                        "工作簿中已存在“" + options.WorksheetName + "”工作表，请删除后重试。");
                }

                if (workbook.ProtectStructure)
                {
                    throw new HostOperationException("工作簿结构已保护，请先取消保护后重试。");
                }

                // 先完成所有读取和内容准备，再删除旧表、创建新表，避免前置校验失败时改动工作簿。
                var visibleWorksheetNames = ReadVisibleWorksheetNames(workbook, existingWorksheet);
                var title = BuildTitle(workbook.Name, options);

                using (CreateStateScope(new HostStateOptions
                {
                    DisableScreenUpdating = true,
                    DisableEvents = true,
                    DisableDisplayAlerts = true
                }))
                {
                    // 先创建并完整写入临时工作表。旧目录在此之前不做任何破坏性修改，
                    // 即使 COM 写入、样式或超链接失败，也只需删除临时表即可恢复原状。
                    var stagingName = BuildTemporaryWorksheetName(workbook);
                    var firstSheet = workbook.Sheets[1];
                    directoryWorksheet = workbook.Worksheets.Add(
                        firstSheet,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing) as Excel.Worksheet;

                    if (directoryWorksheet == null)
                    {
                        throw new HostOperationException("无法创建“" + options.WorksheetName + "”工作表。");
                    }

                    directoryWorksheet.Name = stagingName;
                    WriteDirectoryContents(directoryWorksheet, title, visibleWorksheetNames, options);

                    if (existingWorksheet != null)
                    {
                        originalWorksheetName = SafeWorksheetName(existingWorksheet);
                        var backupName = BuildTemporaryWorksheetName(workbook);

                        // 改名而非立即删除，确保在新目录准备完成后仍有可恢复的旧表。
                        existingRenamed = true;
                        existingWorksheet.Name = backupName;
                    }

                    // 旧目录已改名或不存在，此时临时表可以安全取得正式名称。
                    directoryWorksheet.Name = options.WorksheetName;

                    if (existingWorksheet != null)
                    {
                        // 删除动作放在所有新内容写入成功之后。删除失败时不清理任何一份
                        // 已完成的数据，保留“新目录 + 旧目录备份”供用户恢复。
                        existingDeleteAttempted = true;
                        existingWorksheet.Delete();
                    }

                    completed = true;
                }

                return visibleWorksheetNames.Count;
            }
            catch (HostOperationException)
            {
                if (!completed && !existingDeleteAttempted)
                {
                    TryDeleteWorksheet(directoryWorksheet);
                    if (existingRenamed)
                    {
                        TryRenameWorksheet(existingWorksheet, originalWorksheetName);
                    }
                }

                throw;
            }
            catch (Exception ex)
            {
                if (!completed && !existingDeleteAttempted)
                {
                    TryDeleteWorksheet(directoryWorksheet);
                    if (existingRenamed)
                    {
                        TryRenameWorksheet(existingWorksheet, originalWorksheetName);
                    }
                }

                throw new HostOperationException("生成目录失败：" + ex.Message, ex);
            }
        }

        private static void ValidateOptions(DirectoryOptions options)
        {
            if (string.IsNullOrWhiteSpace(options.WorksheetName))
            {
                throw new HostOperationException("目录工作表名称不能为空。");
            }

            if (options.WorksheetName.Length > 31
                || options.WorksheetName.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0)
            {
                throw new HostOperationException("目录工作表名称不符合 Excel/WPS 命名规则。");
            }

            if (options.Headers == null || options.Headers.Length != ColumnCount
                || options.ColumnWidths == null || options.ColumnWidths.Length != ColumnCount)
            {
                throw new HostOperationException("目录模板必须提供 " + ColumnCount + " 列的表头与列宽。");
            }

            for (var i = 0; i < options.ColumnWidths.Length; i++)
            {
                if (double.IsNaN(options.ColumnWidths[i])
                    || double.IsInfinity(options.ColumnWidths[i])
                    || options.ColumnWidths[i] <= 0d)
                {
                    throw new HostOperationException("目录列宽必须是大于 0 的有效数字。");
                }
            }

            if (options.TitleRowHeight <= 0d || options.HeaderRowHeight <= 0d || options.BodyRowHeight <= 0d
                || options.TitleFontSize <= 0 || options.HeaderFontSize <= 0)
            {
                throw new HostOperationException("目录行高和字号必须是大于 0 的有效数字。");
            }

            // 在任何工作表改名、创建或删除之前验证全部颜色，避免参数错误触发破坏性替换。
            ExcelComHelper.ParseOleColor(options.TitleBackgroundColor);
            ExcelComHelper.ParseOleColor(options.HeaderBackgroundColor);
            ExcelComHelper.ParseOleColor(options.AlternateRowColor);
            ExcelComHelper.ParseOleColor(options.BorderColor);
            ExcelComHelper.ParseOleColor(options.TitleTextColor);
            ExcelComHelper.ParseOleColor(options.BodyTextColor);
            ExcelComHelper.ParseOleColor(options.HyperlinkColor);
        }

        private static string BuildTemporaryWorksheetName(Excel.Workbook workbook)
        {
            const string prefix = "__AccuX_";
            const int maxWorksheetNameLength = 31;
            var suffixLength = maxWorksheetNameLength - prefix.Length;

            for (var attempt = 0; attempt < 5; attempt++)
            {
                var suffix = Guid.NewGuid().ToString("N").Substring(0, suffixLength);
                var candidate = prefix + suffix;
                if (FindWorksheet(workbook, candidate) == null)
                {
                    return candidate;
                }
            }

            throw new HostOperationException("无法为目录生成临时工作表名称。");
        }

        private static Excel.Worksheet FindWorksheet(Excel.Workbook workbook, string worksheetName)
        {
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                if (string.Equals(worksheet.Name, worksheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return worksheet;
                }
            }

            return null;
        }

        private static List<string> ReadVisibleWorksheetNames(
            Excel.Workbook workbook,
            Excel.Worksheet excludedWorksheet)
        {
            var names = new List<string>();

            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                // 替换模式下旧目录表会被删除，不应出现在新目录中。
                if (worksheet == excludedWorksheet)
                {
                    continue;
                }

                if (worksheet.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                {
                    names.Add(worksheet.Name ?? string.Empty);
                }
            }

            return names;
        }

        private static string BuildTitle(string workbookName, DirectoryOptions options)
        {
            var title = Path.GetFileNameWithoutExtension(workbookName ?? string.Empty);
            return string.IsNullOrEmpty(title) ? options.FallbackTitle : title + options.TitleSuffix;
        }

        private static void WriteDirectoryContents(
            Excel.Worksheet directoryWorksheet,
            string title,
            IReadOnlyList<string> worksheetNames,
            DirectoryOptions options)
        {
            var titleRange = directoryWorksheet.Range["A1:C1"];
            titleRange.Merge(Type.Missing);
            titleRange.Value2 = title;

            var headerRange = directoryWorksheet.Range["A2:C2"];
            headerRange.Value2 = new object[,]
            {
                { options.Headers[0], options.Headers[1], options.Headers[2] }
            };

            Excel.Range dataRange = null;
            if (worksheetNames.Count == 0)
            {
                ApplyDirectoryStyle(directoryWorksheet, titleRange, headerRange, null, 0, options);
                return;
            }

            var rows = new object[worksheetNames.Count, ColumnCount];
            for (var index = 0; index < worksheetNames.Count; index++)
            {
                rows[index, 0] = index + 1;
                rows[index, 1] = worksheetNames[index];
                rows[index, 2] = string.Empty;
            }

            var firstDataCell = (Excel.Range)directoryWorksheet.Cells[3, 1];
            var lastDataCell = (Excel.Range)directoryWorksheet.Cells[worksheetNames.Count + 2, ColumnCount];
            dataRange = directoryWorksheet.Range[firstDataCell, lastDataCell];
            dataRange.Value2 = rows;

            for (var index = 0; index < worksheetNames.Count; index++)
            {
                var nameCell = (Excel.Range)directoryWorksheet.Cells[index + 3, 2];
                directoryWorksheet.Hyperlinks.Add(
                    nameCell,
                    string.Empty,
                    BuildWorksheetSubAddress(worksheetNames[index]),
                    Type.Missing,
                    worksheetNames[index]);
            }

            ApplyDirectoryStyle(directoryWorksheet, titleRange, headerRange, dataRange, worksheetNames.Count, options);
        }

        private static void ApplyDirectoryStyle(
            Excel.Worksheet directoryWorksheet,
            Excel.Range titleRange,
            Excel.Range headerRange,
            Excel.Range dataRange,
            int worksheetCount,
            DirectoryOptions options)
        {
            var titleBackground = ExcelComHelper.ParseOleColor(options.TitleBackgroundColor);
            var headerBackground = ExcelComHelper.ParseOleColor(options.HeaderBackgroundColor);
            var alternateRowBackground = ExcelComHelper.ParseOleColor(options.AlternateRowColor);
            var borderColor = ExcelComHelper.ParseOleColor(options.BorderColor);
            var titleTextColor = ExcelComHelper.ParseOleColor(options.TitleTextColor);
            var bodyTextColor = ExcelComHelper.ParseOleColor(options.BodyTextColor);
            var hyperlinkColor = ExcelComHelper.ParseOleColor(options.HyperlinkColor);

            titleRange.Interior.Color = titleBackground;
            titleRange.Font.Color = titleTextColor;
            titleRange.Font.Bold = true;
            titleRange.Font.Size = options.TitleFontSize;
            titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            titleRange.WrapText = false;
            titleRange.RowHeight = options.TitleRowHeight;

            headerRange.Interior.Color = headerBackground;
            headerRange.Font.Color = titleTextColor;
            headerRange.Font.Bold = true;
            headerRange.Font.Size = options.HeaderFontSize;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerRange.WrapText = false;
            headerRange.RowHeight = options.HeaderRowHeight;

            var tableRange = worksheetCount > 0
                ? directoryWorksheet.Range["A2:C" + (worksheetCount + 2)]
                : headerRange;
            tableRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            tableRange.Borders.Color = borderColor;
            tableRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            directoryWorksheet.Range["A:A"].ColumnWidth = options.ColumnWidths[0];
            directoryWorksheet.Range["B:B"].ColumnWidth = options.ColumnWidths[1];
            directoryWorksheet.Range["C:C"].ColumnWidth = options.ColumnWidths[2];

            if (dataRange == null)
            {
                return;
            }

            dataRange.Font.Color = bodyTextColor;
            dataRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            dataRange.RowHeight = options.BodyRowHeight;

            var numberRange = directoryWorksheet.Range["A3:A" + (worksheetCount + 2)];
            numberRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            var textRange = directoryWorksheet.Range["B3:C" + (worksheetCount + 2)];
            textRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            var nameRange = directoryWorksheet.Range["B3:B" + (worksheetCount + 2)];
            nameRange.Font.Color = hyperlinkColor;
            nameRange.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle;

            // 隔行底纹属于渲染规则，颜色由业务模板提供。
            for (var index = 0; index < worksheetCount; index += 2)
            {
                var row = index + 3;
                directoryWorksheet.Range["A" + row + ":C" + row].Interior.Color = alternateRowBackground;
            }
        }

        private static string BuildWorksheetSubAddress(string worksheetName)
        {
            var escapedName = (worksheetName ?? string.Empty).Replace("'", "''");
            return "'" + escapedName + "'!A1";
        }

        private void TryDeleteWorksheet(Excel.Worksheet worksheet)
        {
            if (worksheet == null)
            {
                return;
            }

            try
            {
                using (CreateStateScope(new HostStateOptions
                {
                    DisableDisplayAlerts = true
                }))
                {
                    worksheet.Delete();
                }
            }
            catch
            {
                // 生成失败时的清理也不能掩盖原始异常。
            }
        }

        private static void TryRenameWorksheet(Excel.Worksheet worksheet, string name)
        {
            if (worksheet == null || string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            try
            {
                worksheet.Name = name;
            }
            catch
            {
                // 回滚阶段不能掩盖原始异常；调用方仍保留临时表以避免数据丢失。
            }
        }
    }
}
