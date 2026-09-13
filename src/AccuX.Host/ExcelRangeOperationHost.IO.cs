using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using AccuX.Core.Cells;
using AccuX.Core.Operations;
using Excel = Microsoft.Office.Interop.Excel;

namespace AccuX.Host
{
    public sealed partial class ExcelRangeOperationHost
    {
        /// <summary>
        /// 按 RangeTarget 批量读取 Value 与 Formula，归一化为 CLR 数据（规格 §5.2 / §16）。
        /// 整块读取，禁止逐 Cell COM 访问。
        /// </summary>
        public RangeReadResult Read(RangeTarget target)
        {
            var worksheet = ResolveWorksheetOrThrow(target);
            var workbook = ResolveWorkbook(target);
            var isDate1904 = false;
            try
            {
                isDate1904 = workbook != null && workbook.Date1904;
            }
            catch
            {
                // 兼容宿主取不到日期系统时按默认 1900 系统处理。
            }

            Excel.Range range;
            try
            {
                range = worksheet.Range[target.Address];
            }
            catch (Exception ex)
            {
                throw new HostOperationException("无法解析原选区地址：" + ex.Message, ex);
            }

            var rowCount = target.RowCount;
            var columnCount = target.ColumnCount;

            // 隐藏状态按行/列读取一次，随后作为 CLR 元数据传入公共 Pipeline。
            // 不使用 xlCellTypeVisible：该 API 会把可见单元格拆成多区域，且无法保留行/列维度信息。
            var hiddenRows = ReadHiddenRows(worksheet, range, rowCount);
            var hiddenColumns = ReadHiddenColumns(worksheet, range, columnCount);

            // 整块批量读取：一次 Value / 一次 Formula / 一次 NumberFormat。
            // Value2 保留原始浮点序列值，不让 COM 的 Currency/Date Variant 在读取时提前改变金额精度。
            var values = ReadValueMatrix(range, rowCount, columnCount);
            var formulas = ReadFormulaMatrix(range, target);
            var formulaMask = ReadFormulaMask(range, rowCount, columnCount);
            var formats = ReadNumberFormatMatrix(range, rowCount, columnCount);
            var merged = ReadMergedMatrix(range, rowCount, columnCount);

            if (ContainsTrue(merged))
            {
                throw new HostOperationException("V1 不支持包含合并单元格的选区，请取消合并后重试。");
            }

            var cells = new List<CellData>(rowCount * columnCount);

            for (var row = 0; row < rowCount; row++)
            {
                for (var column = 0; column < columnCount; column++)
                {
                    var rawValue = values[row, column];
                    var hasFormula = formulaMask[row, column];
                    var rawFormula = hasFormula ? formulas[row, column]?.ToString() : null;
                    var numberFormat = formats[row, column];
                    var cell = new CellData(row, column)
                    {
                        NumberFormat = numberFormat,
                        IsMerged = merged[row, column],
                        IsHiddenRow = hiddenRows[row],
                        IsHiddenColumn = hiddenColumns[column],
                        Writable = true
                    };

                    var formulaInfo = hasFormula ? BuildFormulaInfo(rawFormula) : null;
                    cell.Formula = formulaInfo;

                    var normalizedValue = NormalizeValue(rawValue);
                    if (IsDateFormat(numberFormat) && normalizedValue is decimal dateSerial)
                    {
                        try
                        {
                            normalizedValue = DateTime.FromOADate((double)dateSerial + (isDate1904 ? 1462d : 0d));
                        }
                        catch
                        {
                            // 不可转换时保留序列值，分类仍会将其标为日期。
                        }
                    }

                    var input = new CellClassificationInput
                    {
                        HasFormula = hasFormula,
                        IsBlank = IsBlankValue(rawValue),
                        // Value2 对日期返回序列值；公式日期也必须根据 NumberFormat 排除出金额处理。
                        IsDate = IsDateFormat(numberFormat),
                        IsError = IsErrorValue(rawValue),
                        Value = normalizedValue
                    };

                    // 公式结果若为数值但格式为日期，需按日期处理。
                    if (hasFormula && !input.IsError && input.Value is DateTime)
                    {
                        input.IsDate = true;
                    }

                    cell.CellType = CellValueClassifier.Classify(input);
                    cell.Value = input.Value;
                    cells.Add(cell);
                }
            }

            return new RangeReadResult(target, cells);
        }

        private static void Fill(bool[,] values, bool value)
        {
            for (var row = 0; row < values.GetLength(0); row++)
            {
                for (var column = 0; column < values.GetLength(1); column++)
                {
                    values[row, column] = value;
                }
            }
        }

        /// <summary>
        /// 按 RangeWritePlan 批量写回同一个 RangeTarget（规格 §11 / §16）。
        /// 数值与公式分别整块写入，禁止逐 Cell 写入。
        /// </summary>
        public void Write(RangeTarget target, RangeWritePlan writePlan)
        {
            if (writePlan == null || writePlan.IsEmpty)
            {
                return;
            }

            var worksheet = ResolveWorksheetOrThrow(target);

            Excel.Range range;
            try
            {
                range = worksheet.Range[target.Address];
            }
            catch (Exception ex)
            {
                throw new HostOperationException("写回时无法解析原选区地址：" + ex.Message, ex);
            }

            var rowCount = target.RowCount;
            var columnCount = target.ColumnCount;

            // 数值与公式分开成批写入：避免把公式和数值混在同一次赋值里。
            var valueMatrix = new object[rowCount, columnCount];
            var valueMask = new bool[rowCount, columnCount];
            var formulaMatrix = new object[rowCount, columnCount];
            var formulaMask = new bool[rowCount, columnCount];
            var formatMatrix = new object[rowCount, columnCount];
            var formatMask = new bool[rowCount, columnCount];

            foreach (var write in writePlan.Writes)
            {
                if (!IsInside(write.Row, write.Column, rowCount, columnCount))
                {
                    continue;
                }

                if (write.IsFormulaWrite)
                {
                    formulaMatrix[write.Row, write.Column] = NormalizeFormulaForWrite(write.Formula);
                    formulaMask[write.Row, write.Column] = true;
                }
                else
                {
                    valueMatrix[write.Row, write.Column] = write.Value;
                    valueMask[write.Row, write.Column] = true;
                }

                if (!string.IsNullOrEmpty(write.NumberFormat))
                {
                    formatMatrix[write.Row, write.Column] = write.NumberFormat;
                    formatMask[write.Row, write.Column] = true;
                }
            }

            // 公式优先写回，随后写数值，保证公式单元格保持公式属性。
            WriteFormulaBlock(range, formulaMatrix, formulaMask, rowCount, columnCount);
            WriteValueBlock(range, valueMatrix, valueMask, rowCount, columnCount);
            WriteFormatBlock(range, formatMatrix, formatMask, rowCount, columnCount);
        }

        /// <summary>
        /// 写回前检查待写单元格是否属于数组公式或动态数组区域。
        /// 这项检查必须发生在任何批量写入之前，否则数组区域可能先改写部分单元格后才失败。
        /// </summary>
        private static WriteCheckResult ValidateSpecialFormulaAreas(
            Excel.Range range,
            RangeWritePlan writePlan,
            int rowCount,
            int columnCount)
        {
            foreach (var write in writePlan.Writes)
            {
                if (write == null || !IsInside(write.Row, write.Column, rowCount, columnCount))
                {
                    continue;
                }

                Excel.Range cell;
                try
                {
                    cell = (Excel.Range)range.Cells[write.Row + 1, write.Column + 1];
                }
                catch (Exception ex)
                {
                    return WriteCheckResult.Failure("无法读取待写单元格的数组状态：" + ex.Message);
                }

                object hasArray;
                try
                {
                    hasArray = cell.HasArray;
                }
                catch (Exception ex)
                {
                    return WriteCheckResult.Failure("无法确认待写单元格是否属于数组公式：" + ex.Message);
                }

                if (!TryConvertBooleanValue(hasArray, out var isArray))
                {
                    return WriteCheckResult.Failure("无法确认待写单元格是否属于数组公式，请重新执行。");
                }

                if (isArray)
                {
                    return WriteCheckResult.Failure("目标区域包含数组公式，无法安全写回；请先取消数组公式后重试。");
                }

                if (!TryReadDynamicArrayState(cell, out var isDynamicArray, out var issue))
                {
                    return WriteCheckResult.Failure(issue ?? "无法确认待写单元格的动态数组状态，请重新执行。");
                }

                if (isDynamicArray)
                {
                    return WriteCheckResult.Failure("目标区域包含动态数组，无法安全写回；请先取消动态数组后重试。");
                }
            }

            return WriteCheckResult.Success();
        }

        /// <summary>
        /// 读取动态数组状态。PIA 15 没有 HasSpill/SpillParent 属性，因此优先通过 IDispatch
        /// late binding 读取；旧宿主不提供这些属性时退回到公式标记识别。
        /// </summary>
        private static bool TryReadDynamicArrayState(Excel.Range cell, out bool isDynamicArray, out string issue)
        {
            isDynamicArray = false;
            issue = null;

            if (TryReadLateBoundProperty(cell, "HasSpill", out var hasSpill))
            {
                if (!TryConvertBooleanValue(hasSpill, out isDynamicArray))
                {
                    issue = "无法解析宿主返回的动态数组状态。";
                    return false;
                }

                return true;
            }

            if (TryReadLateBoundProperty(cell, "SpillParent", out var spillParent))
            {
                isDynamicArray = spillParent != null && !(spillParent is DBNull);
                return true;
            }

            object rawFormula;
            try
            {
                rawFormula = cell.Formula;
            }
            catch (Exception ex)
            {
                issue = "无法读取待写单元格公式以确认动态数组状态：" + ex.Message;
                return false;
            }

            isDynamicArray = ContainsDynamicArrayMarker(rawFormula?.ToString());
            return true;
        }

        private static bool TryReadLateBoundProperty(Excel.Range range, string propertyName, out object value)
        {
            value = null;
            if (range == null)
            {
                return false;
            }

            try
            {
                value = range.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    range,
                    null,
                    CultureInfo.InvariantCulture);
                return true;
            }
            catch
            {
                // 当前 PIA/宿主不提供该动态数组属性时，交给公式标记回退逻辑处理。
                return false;
            }
        }

        // ---------- 批量读取辅助 ----------

        private static object[,] ReadValueMatrix(Excel.Range range, int rowCount, int columnCount)
        {
            // COM RCW 的 GetType() 通常返回 System.__ComObject，反射取属性不可靠，必须直接访问。
            object raw = range.Value2;
            return ToMatrix(raw, rowCount, columnCount);
        }

        private object[,] ReadFormulaMatrix(Excel.Range range, RangeTarget target)
        {
            object raw;
            try
            {
                raw = range.Formula;
            }
            catch (Exception ex)
            {
                // 公式读取失败不能退化为“无公式”，否则后续可能把公式计算结果当成常量写回。
                throw new HostOperationException("无法读取选区公式，请重试。", ex);
            }

            return ToMatrix(raw, target.RowCount, target.ColumnCount);
        }

        /// <summary>
        /// 批量识别真正的公式单元格（规格 §8 Formula 处理契约）。
        /// <para>
        /// 不能用 <c>range.Formula</c> 是否为空来判定：常量单元格的 Formula 属性返回的是字面值
        /// （数值返回 "12345"、文本返回 "text"），据此判定会把所有常量误判为公式。
        /// 也不能用 <c>range.HasFormula</c>：选区混合时它返回 DBNull，无法逐格区分。
        /// 因此用 <c>SpecialCells(xlCellTypeFormulas)</c> 取得公式区域，再映射回 0 基掩码。
        /// </para>
        /// </summary>
        private static bool[,] ReadFormulaMask(Excel.Range range, int rowCount, int columnCount)
        {
            var result = new bool[rowCount, columnCount];
            object hasFormula;

            try
            {
                hasFormula = range.HasFormula;
            }
            catch (Exception ex)
            {
                throw new HostOperationException("无法确认选区公式状态，请重试。", ex);
            }

            // HasFormula 为明确的布尔值时，False 表示整个选区没有公式，True 表示整个选区都是公式。
            // 混合区域通常返回 Null/DBNull，此时再用 SpecialCells 映射每个公式区域。
            if (TryConvertBooleanValue(hasFormula, out var allFormula))
            {
                if (allFormula)
                {
                    Fill(result, true);
                }

                return result;
            }

            Excel.Range formulaRange;
            try
            {
                formulaRange = range.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
            }
            catch (Exception ex)
            {
                // 此处已经确认选区不是明确的“无公式”状态；失败必须终止，不能按无公式继续写回。
                throw new HostOperationException("无法确认选区中的公式单元格，请重试。", ex);
            }

            if (formulaRange == null)
            {
                throw new HostOperationException("宿主未返回选区公式区域，无法安全写回。");
            }

            try
            {
                var baseRow = range.Row;
                var baseColumn = range.Column;

                foreach (Excel.Range area in formulaRange.Areas)
                {
                    MarkArea(
                        result,
                        area.Row - baseRow,
                        area.Column - baseColumn,
                        SafeRowCount(area),
                        SafeColumnCount(area));
                }
            }
            catch (Exception ex)
            {
                throw new HostOperationException("无法映射选区公式区域，请重试。", ex);
            }

            return result;
        }

        /// <summary>
        /// 将一块矩形区域标记进掩码（纯逻辑，可脱离 COM 单元测试）。
        /// 越界部分自动裁剪。
        /// </summary>
        internal static void MarkArea(bool[,] mask, int startRow, int startColumn, int areaRows, int areaColumns)
        {
            if (mask == null)
            {
                return;
            }

            var rowCount = mask.GetLength(0);
            var columnCount = mask.GetLength(1);

            for (var row = 0; row < areaRows; row++)
            {
                var targetRow = startRow + row;
                if (targetRow < 0 || targetRow >= rowCount)
                {
                    continue;
                }

                for (var column = 0; column < areaColumns; column++)
                {
                    var targetColumn = startColumn + column;
                    if (targetColumn < 0 || targetColumn >= columnCount)
                    {
                        continue;
                    }

                    mask[targetRow, targetColumn] = true;
                }
            }
        }

        private static string[,] ReadNumberFormatMatrix(Excel.Range range, int rowCount, int columnCount)
        {
            var result = new string[rowCount, columnCount];
            try
            {
                var raw = ToMatrix(range.NumberFormat, rowCount, columnCount);
                for (var row = 0; row < rowCount; row++)
                {
                    for (var column = 0; column < columnCount; column++)
                    {
                        var value = raw[row, column];
                        result[row, column] = value?.ToString() ?? string.Empty;
                    }
                }
            }
            catch (Exception ex)
            {
                // NumberFormat 参与日期识别。读取失败时不能退化为普通数值，
                // 否则日期序列可能被财务转换写回。
                throw new HostOperationException("无法读取选区数字格式，请重试。", ex);
            }

            return result;
        }

        private static bool[,] ReadMergedMatrix(Excel.Range range, int rowCount, int columnCount)
        {
            var result = new bool[rowCount, columnCount];
            try
            {
                var raw = ToMatrix(range.MergeCells, rowCount, columnCount);
                for (var row = 0; row < rowCount; row++)
                {
                    for (var column = 0; column < columnCount; column++)
                    {
                        // 混合状态在不同宿主中可能以 null/DBNull 或矩阵中的非 bool
                        // 表示；未知状态一律按“可能包含合并”处理。
                        if (raw[row, column] is bool merged)
                        {
                            result[row, column] = merged;
                        }
                        else
                        {
                            result[row, column] = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new HostOperationException("无法读取选区合并单元格状态，请重试。", ex);
            }

            return result;
        }

        private static bool ContainsTrue(bool[,] matrix)
        {
            if (matrix == null)
            {
                return true;
            }

            for (var row = 0; row < matrix.GetLength(0); row++)
            {
                for (var column = 0; column < matrix.GetLength(1); column++)
                {
                    if (matrix[row, column])
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// 将 COM 返回值转换为 0 基、尺寸等于目标区域的 2D 矩阵。
        /// <para>
        /// Excel/WPS 的 COM 数组下界固定为 1（<c>GetLowerBound(0) == 1</c>），即使尺寸与目标一致也不能直接返回，
        /// 否则调用方按 0 基索引访问会抛 IndexOutOfRangeException。
        /// 单单元格或整块同值时 COM 会返回标量，此处按目标尺寸广播。
        /// </para>
        /// </summary>
        internal static object[,] ToMatrix(object raw, int rowCount = 1, int columnCount = 1)
        {
            if (rowCount < 0)
            {
                rowCount = 0;
            }

            if (columnCount < 0)
            {
                columnCount = 0;
            }

            var result = new object[rowCount, columnCount];

            if (raw is object[,] matrix)
            {
                // 按源数组实际下界取值，兼容 0 基与 1 基；尺寸不一致时只复制重叠区域。
                var lowerRow = matrix.GetLowerBound(0);
                var lowerColumn = matrix.GetLowerBound(1);
                var rows = Math.Min(rowCount, matrix.GetLength(0));
                var columns = Math.Min(columnCount, matrix.GetLength(1));

                for (var row = 0; row < rows; row++)
                {
                    for (var column = 0; column < columns; column++)
                    {
                        result[row, column] = matrix[lowerRow + row, lowerColumn + column];
                    }
                }

                return result;
            }

            // 标量：广播到整个目标区域。
            for (var row = 0; row < rowCount; row++)
            {
                for (var column = 0; column < columnCount; column++)
                {
                    result[row, column] = raw;
                }
            }

            return result;
        }

        // ---------- 值归一化 ----------

        private static bool IsBlankValue(object value)
        {
            return value == null || value is DBNull;
        }

        private static bool IsErrorValue(object value)
        {
            return value is int errorCode && IsExcelError(errorCode);
        }

        private static bool IsExcelError(int code)
        {
            switch (code)
            {
                case -2146826281:
                case -2146826246:
                case -2146826259:
                case -2146826288:
                case -2146826252:
                case -2146826265:
                case -2146826273:
                    return true;
                default:
                    return false;
            }
        }

        private static object NormalizeValue(object raw)
        {
            if (raw == null || raw is DBNull)
            {
                return null;
            }

            if (raw is bool)
            {
                return raw;
            }

            if (IsErrorValue(raw))
            {
                return raw;
            }

            if (raw is DateTime)
            {
                return raw;
            }

            if (raw is string text)
            {
                return text;
            }

            if (IsNumeric(raw))
            {
                return CellValueClassifier.NormalizeNumeric(raw);
            }

            return raw;
        }

        private static bool IsNumeric(object value)
        {
            switch (value)
            {
                case decimal _:
                case double _:
                case float _:
                case int _:
                case long _:
                case short _:
                case byte _:
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// 判断 NumberFormat 是否包含日期/时间占位符。
        /// <para>
        /// Excel 自定义格式中的方括号可能表示颜色（例如 <c>[Red]</c>）、条件
        /// （例如 <c>[&gt;=100]</c>）或区域设置（例如 <c>[$-409]</c>），引号、反斜杠、
        /// 下划线和星号后面的字符也可能只是字面量/填充字符，不能直接按字符搜索。
        /// </para>
        /// </summary>
        internal static bool IsDateFormat(string numberFormat)
        {
            if (string.IsNullOrEmpty(numberFormat))
            {
                return false;
            }

            for (var index = 0; index < numberFormat.Length; index++)
            {
                var current = numberFormat[index];

                // 引号中的内容是字面量；Excel 允许用连续双引号表示字面量引号。
                if (current == '"')
                {
                    if (index + 1 < numberFormat.Length && numberFormat[index + 1] == '"')
                    {
                        index++;
                        continue;
                    }

                    var quoteEnd = numberFormat.IndexOf('"', index + 1);
                    if (quoteEnd < 0)
                    {
                        break;
                    }

                    index = quoteEnd;
                    continue;
                }

                // 反斜杠、下划线和星号分别转义/占用后面的一个字符。
                if (current == '\\' || current == '_' || current == '*')
                {
                    if (index + 1 < numberFormat.Length)
                    {
                        index++;
                    }

                    continue;
                }

                // 方括号内容优先按控制段处理，避免 [Red] 中的 d 被判为日期。
                if (current == '[')
                {
                    var bracketEnd = numberFormat.IndexOf(']', index + 1);
                    if (bracketEnd >= 0)
                    {
                        var bracketToken = numberFormat.Substring(index + 1, bracketEnd - index - 1);
                        if (IsElapsedTimeToken(bracketToken))
                        {
                            return true;
                        }

                        index = bracketEnd;
                        continue;
                    }
                }

                var token = char.ToLowerInvariant(current);
                if (token == 'y' || token == 'd' || token == 'h' || token == 's')
                {
                    return true;
                }

                if (token == 'm')
                {
                    var tokenStart = index;
                    while (index + 1 < numberFormat.Length
                        && char.ToLowerInvariant(numberFormat[index + 1]) == 'm')
                    {
                        index++;
                    }

                    var tokenLength = index - tokenStart + 1;
                    if (tokenLength >= 3
                        || IsStandaloneMonthToken(numberFormat, tokenStart, index)
                        || HasDateSeparator(numberFormat, tokenStart, index))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool IsElapsedTimeToken(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return false;
            }

            var normalized = token.Trim().ToLowerInvariant();
            return normalized == "h"
                || normalized == "hh"
                || normalized == "m"
                || normalized == "mm"
                || normalized == "s"
                || normalized == "ss";
        }

        private static bool IsStandaloneMonthToken(string format, int tokenStart, int tokenEnd)
        {
            var tokenLength = tokenEnd - tokenStart + 1;
            if (tokenLength > 2)
            {
                return true;
            }

            for (var index = 0; index < format.Length; index++)
            {
                if (index >= tokenStart && index <= tokenEnd)
                {
                    index = tokenEnd;
                    continue;
                }

                var current = format[index];
                if (char.IsWhiteSpace(current) || current == ';')
                {
                    continue;
                }

                return false;
            }

            return true;
        }

        private static bool HasDateSeparator(string format, int tokenStart, int tokenEnd)
        {
            var before = tokenStart > 0 ? format[tokenStart - 1] : '\0';
            var after = tokenEnd + 1 < format.Length ? format[tokenEnd + 1] : '\0';
            return before == '/' || before == '-' || before == ':' || before == '.'
                || after == '/' || after == '-' || after == ':' || after == '.';
        }

        // ---------- 公式规范化 ----------

        private FormulaInfo BuildFormulaInfo(string rawFormula)
        {
            var expression = NormalizeFormulaForRead(rawFormula);
            var kind = DetectFormulaKind(rawFormula, expression);
            var canTransform = kind == FormulaKind.Normal;

            return new FormulaInfo(expression, kind, canTransform);
        }

        /// <summary>
        /// 将宿主公式表达式规范化为统一形式（统一 '=' 前缀）。
        /// 区域设置导致的参数分隔符差异在 V1 中暂不做替换：Excel/WPS 中文环境均使用逗号。
        /// </summary>
        private static string NormalizeFormulaForRead(string rawFormula)
        {
            if (string.IsNullOrEmpty(rawFormula))
            {
                return string.Empty;
            }

            var text = rawFormula.Trim();
            return text.StartsWith("=", StringComparison.Ordinal) ? text : "=" + text;
        }

        private static FormulaKind DetectFormulaKind(string rawFormula, string normalized)
        {
            if (string.IsNullOrEmpty(normalized))
            {
                return FormulaKind.Unsupported;
            }

            // 数组公式在 Formula 属性中通常带 { } 包裹。
            if (rawFormula != null && rawFormula.TrimStart().StartsWith("{", StringComparison.Ordinal))
            {
                return FormulaKind.Array;
            }

            // 动态数组/溢出公式特征：动态函数标记或真正的溢出引用（例如 A1#）。
            // 不能只搜索任意 '#'，否则 Table1[#Data]、"#" 和 #N/A 都会被误判。
            if (ContainsDynamicArrayMarker(normalized))
            {
                return FormulaKind.DynamicArray;
            }

            return FormulaKind.Normal;
        }

        private static bool ContainsDynamicArrayMarker(string formula)
        {
            if (string.IsNullOrEmpty(formula))
            {
                return false;
            }

            if (formula.IndexOf("_xlfn.", StringComparison.OrdinalIgnoreCase) >= 0 ||
                formula.IndexOf("_xlws.", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            var inString = false;
            for (var index = 0; index < formula.Length; index++)
            {
                var current = formula[index];
                if (current == '"')
                {
                    // Excel 使用两个双引号表示字符串中的一个双引号。
                    if (inString && index + 1 < formula.Length && formula[index + 1] == '"')
                    {
                        index++;
                        continue;
                    }

                    inString = !inString;
                    continue;
                }

                if (inString || current != '#' || index == 0)
                {
                    continue;
                }

                var previous = formula[index - 1];
                if (char.IsLetterOrDigit(previous) || previous == '$' || previous == ')')
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 将业务层生成的规范化公式转换为当前宿主可安全接受的写入表达式。
        /// </summary>
        private static string NormalizeFormulaForWrite(string formula)
        {
            if (string.IsNullOrEmpty(formula))
            {
                return string.Empty;
            }

            var text = formula.Trim();
            return text.StartsWith("=", StringComparison.Ordinal) ? text : "=" + text;
        }

        // ---------- 批量写入辅助 ----------

        private static bool IsInside(int row, int column, int rowCount, int columnCount)
        {
            return row >= 0 && row < rowCount && column >= 0 && column < columnCount;
        }

        private static void WriteValueBlock(Excel.Range range, object[,] matrix, bool[,] mask, int rowCount, int columnCount)
        {
            if (!HasAny(mask, rowCount, columnCount))
            {
                return;
            }

            // 若整个选区都是数值写入，可一次性赋值；否则按行分段批量写入，避免逐 Cell。
            if (IsFull(mask, rowCount, columnCount))
            {
                range.Value2 = matrix;
                return;
            }

            WriteBlockByMask(range, matrix, mask, rowCount, columnCount, (target, block) => target.Value2 = block);
        }

        private static void WriteFormulaBlock(Excel.Range range, object[,] matrix, bool[,] mask, int rowCount, int columnCount)
        {
            if (!HasAny(mask, rowCount, columnCount))
            {
                return;
            }

            if (IsFull(mask, rowCount, columnCount))
            {
                range.Formula = matrix;
                return;
            }

            WriteBlockByMask(range, matrix, mask, rowCount, columnCount, (target, block) => target.Formula = block);
        }

        private static void WriteFormatBlock(Excel.Range range, object[,] matrix, bool[,] mask, int rowCount, int columnCount)
        {
            if (!HasAny(mask, rowCount, columnCount))
            {
                return;
            }

            // NumberFormat 的 COM 属性并不支持像 Value2/Formula 一样稳定地接收二维数组。
            // Excel/WPS 在此处可能已经完成了 Value2 写回，但对二维 NumberFormat 赋值返回
            // 0x80004005 (E_FAIL)，从而表现为“结果已写入但命令报错”。按连续且格式相同的
            // 区段使用标量字符串写回，既保留批量写入，又兼容两种宿主。
            for (var row = 0; row < rowCount; row++)
            {
                var column = 0;
                while (column < columnCount)
                {
                    if (!mask[row, column])
                    {
                        column++;
                        continue;
                    }

                    var start = column;
                    var format = matrix[row, column]?.ToString();
                    column++;

                    while (column < columnCount
                        && mask[row, column]
                        && string.Equals(format, matrix[row, column]?.ToString(), StringComparison.Ordinal))
                    {
                        column++;
                    }

                    var width = column - start;
                    var target = ((Excel.Range)range.Cells[row + 1, start + 1]).Resize[1, width];
                    target.NumberFormat = format;
                }
            }
        }

        /// <summary>
        /// 按行连续区段批量写入，避免逐 Cell COM 调用（规格 §16）。
        /// </summary>
        private static void WriteBlockByMask(
            Excel.Range range,
            object[,] matrix,
            bool[,] mask,
            int rowCount,
            int columnCount,
            Action<Excel.Range, object[,]> write)
        {
            for (var row = 0; row < rowCount; row++)
            {
                var column = 0;
                while (column < columnCount)
                {
                    if (!mask[row, column])
                    {
                        column++;
                        continue;
                    }

                    var start = column;
                    while (column < columnCount && mask[row, column])
                    {
                        column++;
                    }

                    var width = column - start;
                    var block = new object[1, width];
                    for (var offset = 0; offset < width; offset++)
                    {
                        block[0, offset] = matrix[row, start + offset];
                    }

                    var target = ((Excel.Range)range.Cells[row + 1, start + 1]).Resize[1, width];
                    write(target, block);
                }
            }
        }

        private static bool HasAny(bool[,] mask, int rowCount, int columnCount)
        {
            for (var row = 0; row < rowCount; row++)
            {
                for (var column = 0; column < columnCount; column++)
                {
                    if (mask[row, column])
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool IsFull(bool[,] mask, int rowCount, int columnCount)
        {
            for (var row = 0; row < rowCount; row++)
            {
                for (var column = 0; column < columnCount; column++)
                {
                    if (!mask[row, column])
                    {
                        return false;
                    }
                }
            }

            return true;
        }
    }
}
