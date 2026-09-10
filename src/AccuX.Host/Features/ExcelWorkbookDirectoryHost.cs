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
        /// 在当前工作簿最前面生成目录工作表；replaceExisting 为 true 时先删除同名旧表。
        /// </summary>
        public int GenerateDirectory(DirectoryOptions options, bool replaceExisting)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            ValidateOptions(options);

            Excel.Workbook workbook = null;
            Excel.Worksheet directoryWorksheet = null;
            var completed = false;

            try
            {
                workbook = GetActiveWorkbook();
                var existingWorksheet = FindWorksheet(workbook, options.WorksheetName);
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
                    if (existingWorksheet != null)
                    {
                        existingWorksheet.Delete();
                    }

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

                    directoryWorksheet.Name = options.WorksheetName;
                    WriteDirectoryContents(directoryWorksheet, title, visibleWorksheetNames, options);
                    completed = true;
                }

                return visibleWorksheetNames.Count;
            }
            catch (HostOperationException)
            {
                if (!completed)
                {
                    TryDeleteWorksheet(directoryWorksheet);
                }

                throw;
            }
            catch (Exception ex)
            {
                if (!completed)
                {
                    TryDeleteWorksheet(directoryWorksheet);
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

            if (options.Headers == null || options.Headers.Length != ColumnCount
                || options.ColumnWidths == null || options.ColumnWidths.Length != ColumnCount)
            {
                throw new HostOperationException("目录模板必须提供 " + ColumnCount + " 列的表头与列宽。");
            }
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
    }
}
