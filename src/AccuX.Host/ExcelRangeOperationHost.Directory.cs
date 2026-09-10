using System;
using System.Collections.Generic;
using System.IO;
using AccuX.Core.Operations;
using Excel = Microsoft.Office.Interop.Excel;

namespace AccuX.Host
{
    public sealed partial class ExcelRangeOperationHost
    {
        private const string DirectoryWorksheetName = "目录";
        private static readonly int TitleBackgroundColor = Rgb(20, 49, 70);
        private static readonly int HeaderBackgroundColor = Rgb(43, 99, 145);
        private static readonly int AlternateRowBackgroundColor = Rgb(242, 246, 250);
        private static readonly int BorderColor = Rgb(210, 218, 226);
        private static readonly int WhiteColor = Rgb(255, 255, 255);
        private static readonly int BodyTextColor = Rgb(31, 41, 55);
        private static readonly int HyperlinkColor = Rgb(5, 99, 193);

        /// <summary>
        /// 判断当前活动工作簿中是否已存在“目录”工作表。
        /// </summary>
        public bool DirectoryWorksheetExists()
        {
            var workbook = GetActiveWorkbook();
            return FindDirectoryWorksheet(workbook) != null;
        }

        /// <summary>
        /// 在当前工作簿最前面生成工作表目录。
        /// replaceExisting 为 true 时先删除已存在的“目录”工作表再重新生成。
        /// </summary>
        public int GenerateDirectory(bool replaceExisting)
        {
            Excel.Workbook workbook = null;
            Excel.Worksheet directoryWorksheet = null;
            var completed = false;

            try
            {
                workbook = GetActiveWorkbook();
                var existingDirectoryWorksheet = FindDirectoryWorksheet(workbook);
                if (existingDirectoryWorksheet != null && !replaceExisting)
                {
                    throw new HostOperationException("工作簿中已存在“目录”工作表，请删除后重试。");
                }

                if (workbook.ProtectStructure)
                {
                    throw new HostOperationException("工作簿结构已保护，请先取消保护后重试。");
                }

                // 先完成所有读取和内容准备，再删除旧表、创建新表，避免前置校验失败时改动工作簿。
                var visibleWorksheetNames = ReadVisibleWorksheetNames(workbook, existingDirectoryWorksheet);
                var title = BuildDirectoryTitle(workbook.Name);

                using (BeginStateScope(new HostStateOptions
                {
                    DisableScreenUpdating = true,
                    DisableEvents = true,
                    DisableDisplayAlerts = true
                }))
                {
                    if (existingDirectoryWorksheet != null)
                    {
                        existingDirectoryWorksheet.Delete();
                    }

                    var firstSheet = workbook.Sheets[1];
                    directoryWorksheet = workbook.Worksheets.Add(
                        firstSheet,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing) as Excel.Worksheet;

                    if (directoryWorksheet == null)
                    {
                        throw new HostOperationException("无法创建“目录”工作表。");
                    }

                    directoryWorksheet.Name = DirectoryWorksheetName;
                    WriteDirectoryContents(directoryWorksheet, title, visibleWorksheetNames);
                    completed = true;
                }

                return visibleWorksheetNames.Count;
            }
            catch (HostOperationException)
            {
                if (!completed)
                {
                    TryDeleteCreatedDirectoryWorksheet(directoryWorksheet);
                }

                throw;
            }
            catch (Exception ex)
            {
                if (!completed)
                {
                    TryDeleteCreatedDirectoryWorksheet(directoryWorksheet);
                }

                throw new HostOperationException("生成目录失败：" + ex.Message, ex);
            }
        }

        private static Excel.Worksheet FindDirectoryWorksheet(Excel.Workbook workbook)
        {
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                if (string.Equals(worksheet.Name, DirectoryWorksheetName, StringComparison.OrdinalIgnoreCase))
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
                // 替换模式下旧“目录”表会被删除，不应出现在新目录中。
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

        private static string BuildDirectoryTitle(string workbookName)
        {
            var title = Path.GetFileNameWithoutExtension(workbookName ?? string.Empty);
            return string.IsNullOrEmpty(title) ? "工作簿目录" : title + DirectoryWorksheetName;
        }

        private static void WriteDirectoryContents(
            Excel.Worksheet directoryWorksheet,
            string title,
            IReadOnlyList<string> worksheetNames)
        {
            var titleRange = directoryWorksheet.Range["A1:C1"];
            titleRange.Merge(Type.Missing);
            titleRange.Value2 = title;

            var headerRange = directoryWorksheet.Range["A2:C2"];
            headerRange.Value2 = new object[,]
            {
                { "序号", "名称", "备注" }
            };

            Excel.Range dataRange = null;
            if (worksheetNames.Count == 0)
            {
                ApplyDirectoryStyle(directoryWorksheet, titleRange, headerRange, null, 0);
                return;
            }

            var rows = new object[worksheetNames.Count, 3];
            for (var index = 0; index < worksheetNames.Count; index++)
            {
                rows[index, 0] = index + 1;
                rows[index, 1] = worksheetNames[index];
                rows[index, 2] = string.Empty;
            }

            var firstDataCell = (Excel.Range)directoryWorksheet.Cells[3, 1];
            var lastDataCell = (Excel.Range)directoryWorksheet.Cells[worksheetNames.Count + 2, 3];
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

            ApplyDirectoryStyle(directoryWorksheet, titleRange, headerRange, dataRange, worksheetNames.Count);
        }

        private static void ApplyDirectoryStyle(
            Excel.Worksheet directoryWorksheet,
            Excel.Range titleRange,
            Excel.Range headerRange,
            Excel.Range dataRange,
            int worksheetCount)
        {
            titleRange.Interior.Color = TitleBackgroundColor;
            titleRange.Font.Color = WhiteColor;
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 16;
            titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            titleRange.WrapText = false;
            titleRange.RowHeight = 30;

            headerRange.Interior.Color = HeaderBackgroundColor;
            headerRange.Font.Color = WhiteColor;
            headerRange.Font.Bold = true;
            headerRange.Font.Size = 11;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerRange.WrapText = false;
            headerRange.RowHeight = 23;

            var tableRange = worksheetCount > 0
                ? directoryWorksheet.Range["A2:C" + (worksheetCount + 2)]
                : headerRange;
            tableRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            tableRange.Borders.Color = BorderColor;
            tableRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            directoryWorksheet.Range["A:A"].ColumnWidth = 8;
            directoryWorksheet.Range["B:B"].ColumnWidth = 30;
            directoryWorksheet.Range["C:C"].ColumnWidth = 42;

            if (dataRange == null)
            {
                return;
            }

            dataRange.Font.Color = BodyTextColor;
            dataRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            dataRange.RowHeight = 20;

            var numberRange = directoryWorksheet.Range["A3:A" + (worksheetCount + 2)];
            numberRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            var textRange = directoryWorksheet.Range["B3:C" + (worksheetCount + 2)];
            textRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            var nameRange = directoryWorksheet.Range["B3:B" + (worksheetCount + 2)];
            nameRange.Font.Color = HyperlinkColor;
            nameRange.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle;

            for (var index = 0; index < worksheetCount; index += 2)
            {
                var row = index + 3;
                directoryWorksheet.Range["A" + row + ":C" + row].Interior.Color = AlternateRowBackgroundColor;
            }
        }

        private static int Rgb(int red, int green, int blue)
        {
            return red + (green << 8) + (blue << 16);
        }

        private static string BuildWorksheetSubAddress(string worksheetName)
        {
            var escapedName = (worksheetName ?? string.Empty).Replace("'", "''");
            return "'" + escapedName + "'!A1";
        }

        private void TryDeleteCreatedDirectoryWorksheet(Excel.Worksheet worksheet)
        {
            if (worksheet == null)
            {
                return;
            }

            try
            {
                using (BeginStateScope(new HostStateOptions
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
