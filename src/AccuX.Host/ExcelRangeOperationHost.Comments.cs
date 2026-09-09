using System;
using AccuX.Core.Operations;
using Excel = Microsoft.Office.Interop.Excel;

namespace AccuX.Host
{
    public sealed partial class ExcelRangeOperationHost
    {
        /// <summary>
        /// 固定传统批注目标的轻量选择路径。
        /// 这里只遍历 Areas 的首单元格，不读取选区值，也不套用批量处理上限。
        /// </summary>
        public CellCommentTarget CaptureCommentTarget()
        {
            var workbook = GetActiveWorkbook();
            var worksheet = GetActiveWorksheet(workbook);
            var selection = GetSelectionRange(worksheet);

            if (selection == null)
            {
                throw new HostOperationException("当前选区不是有效的单元格区域，请先选择一个数据区域。");
            }

            Excel.Range bestCell = null;
            var areaCount = SafeAreaCount(selection);
            for (var areaIndex = 1; areaIndex <= areaCount; areaIndex++)
            {
                Excel.Range area;
                try
                {
                    area = selection.Areas[areaIndex];
                }
                catch (Exception ex)
                {
                    throw new HostOperationException("无法读取当前选区，请重新选择一个单元格。", ex);
                }

                var candidate = GetFirstCell(area);
                candidate = GetMergedAnchor(candidate);
                if (candidate == null)
                {
                    continue;
                }

                if (bestCell == null || IsEarlierCell(candidate, bestCell))
                {
                    bestCell = candidate;
                }
            }

            if (bestCell == null)
            {
                throw new HostOperationException("当前选区为空，请先选择一个单元格。");
            }

            var address = SafeAddress(bestCell);
            if (string.IsNullOrWhiteSpace(address))
            {
                throw new HostOperationException("无法确定批注目标地址，请重新选择一个单元格。");
            }

            return new CellCommentTarget(
                BuildWorkbookKey(workbook),
                SafeWorksheetName(worksheet),
                SafeWorksheetName(worksheet),
                address,
                SafeCellRow(bestCell),
                SafeCellColumn(bestCell));
        }

        /// <summary>读取目标单元格的传统批注文本。</summary>
        public CellComment ReadComment(CellCommentTarget target)
        {
            var range = ResolveCommentRangeOrThrow(target, false);
            try
            {
                var comment = range.Comment;
                if (comment == null)
                {
                    return CellComment.None;
                }

                var text = comment.Text(Type.Missing, Type.Missing, Type.Missing);
                return new CellComment(true, text ?? string.Empty);
            }
            catch (Exception ex)
            {
                throw new HostOperationException("读取单元格批注失败：" + ex.Message, ex);
            }
        }

        /// <summary>覆盖或新增目标单元格的传统批注。</summary>
        public void SaveComment(CellCommentTarget target, string text)
        {
            if (text == null)
            {
                throw new ArgumentNullException(nameof(text));
            }

            var range = ResolveCommentRangeOrThrow(target, true);
            try
            {
                // 保存始终按“覆盖”处理：清除目标单元格的旧备注后直接创建新备注。
                // ClearComments 在没有旧备注时是无操作，因此无需先读取/判断 Comment 是否存在。
                range.ClearComments();
                range.AddComment(text);
            }
            catch (Exception ex)
            {
                throw new HostOperationException("保存单元格批注失败：" + ex.Message, ex);
            }
        }

        /// <summary>删除目标单元格的传统批注。</summary>
        public void DeleteComment(CellCommentTarget target)
        {
            var range = ResolveCommentRangeOrThrow(target, true);
            try
            {
                var comment = range.Comment;
                if (comment != null)
                {
                    comment.Delete();
                }
            }
            catch (Exception ex)
            {
                throw new HostOperationException("删除单元格批注失败：" + ex.Message, ex);
            }
        }

        private Excel.Range ResolveCommentRangeOrThrow(CellCommentTarget target, bool forWrite)
        {
            if (target == null)
            {
                throw new HostOperationException("批注目标已丢失，请重新执行。");
            }

            Excel.Workbook workbook = null;
            try
            {
                foreach (Excel.Workbook candidate in _application.Workbooks)
                {
                    if (string.Equals(BuildWorkbookKey(candidate), target.WorkbookKey, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = candidate;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new HostOperationException("无法确认原工作簿是否仍处于打开状态。", ex);
            }

            if (workbook == null)
            {
                throw new HostOperationException("原工作簿已关闭，请重新执行。");
            }

            Excel.Worksheet worksheet = null;
            try
            {
                foreach (Excel.Worksheet candidate in workbook.Worksheets)
                {
                    if (string.Equals(SafeWorksheetName(candidate), target.WorksheetKey, StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = candidate;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new HostOperationException("无法确认原工作表是否仍处于打开状态。", ex);
            }

            if (worksheet == null)
            {
                throw new HostOperationException("原工作表已关闭或已重命名，请重新执行。");
            }

            if (forWrite && IsSheetProtected(worksheet))
            {
                throw new HostOperationException("目标工作表处于保护状态，无法修改批注。请先取消保护。");
            }

            try
            {
                var range = worksheet.Range[target.Address];
                if (range == null)
                {
                    throw new HostOperationException("原批注目标地址已失效，请重新执行。");
                }

                if (SafeRowCount(range) != 1 || SafeColumnCount(range) != 1)
                {
                    throw new HostOperationException("批注目标必须是单个单元格，请重新执行。");
                }

                return range;
            }
            catch (HostOperationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new HostOperationException("原批注目标地址已失效：" + ex.Message, ex);
            }
        }

        private static Excel.Range GetFirstCell(Excel.Range area)
        {
            if (area == null)
            {
                return null;
            }

            try
            {
                return area.Cells[1, 1] as Excel.Range;
            }
            catch
            {
                return null;
            }
        }

        private static Excel.Range GetMergedAnchor(Excel.Range cell)
        {
            if (cell == null)
            {
                return null;
            }

            try
            {
                if (cell.MergeCells is bool merged && merged)
                {
                    var mergeArea = cell.MergeArea;
                    return GetFirstCell(mergeArea) ?? cell;
                }
            }
            catch
            {
                // 某些宿主对 MergeCells 访问失败时仍可安全使用原单元格。
            }

            return cell;
        }

        private static bool IsEarlierCell(Excel.Range candidate, Excel.Range current)
        {
            var candidateRow = SafeCellRow(candidate);
            var currentRow = SafeCellRow(current);
            if (candidateRow != currentRow)
            {
                return candidateRow < currentRow;
            }

            return SafeCellColumn(candidate) < SafeCellColumn(current);
        }

        private static int SafeCellRow(Excel.Range cell)
        {
            try
            {
                return Convert.ToInt32(cell.Row);
            }
            catch
            {
                return int.MaxValue;
            }
        }

        private static int SafeCellColumn(Excel.Range cell)
        {
            try
            {
                return Convert.ToInt32(cell.Column);
            }
            catch
            {
                return int.MaxValue;
            }
        }
    }
}
