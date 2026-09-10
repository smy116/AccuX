using System;
using System.Collections.Generic;
using System.Globalization;
using AccuX.Core.Cells;
using AccuX.Core.Operations;
using AccuX.Modules.BasicFinance.ChineseAmount;
using AccuX.Modules.BasicFinance.Rounding;

namespace AccuX.Modules.BasicFinance.SelectionSum
{
    /// <summary>
    /// 选区求和的纯 CLR 结果。
    /// </summary>
    public sealed class SelectionSumResult
    {
        public SelectionSumResult(
            decimal rawTotal,
            decimal roundedTotal,
            string formattedTotal,
            int numericCellCount)
            : this(
                rawTotal,
                roundedTotal,
                formattedTotal,
                (roundedTotal / 10000m).ToString("#,##0.00", CultureInfo.InvariantCulture),
                ChineseAmountService.ToChineseAmount(roundedTotal),
                numericCellCount)
        {
        }

        public SelectionSumResult(
            decimal rawTotal,
            decimal roundedTotal,
            string formattedTotal,
            string formattedWanTotal,
            string chineseTotal,
            int numericCellCount)
        {
            RawTotal = rawTotal;
            RoundedTotal = roundedTotal;
            FormattedTotal = formattedTotal ?? string.Empty;
            FormattedWanTotal = formattedWanTotal ?? string.Empty;
            ChineseTotal = chineseTotal ?? string.Empty;
            NumericCellCount = numericCellCount;
        }

        /// <summary>所有可见数字累加后的未舍入总和。</summary>
        public decimal RawTotal { get; }

        /// <summary>按两位小数、AwayFromZero 舍入后的总和。</summary>
        public decimal RoundedTotal { get; }

        /// <summary>用于“金额”复制项的千分位文本。</summary>
        public string FormattedTotal { get; }

        /// <summary>合计折合为万元后的千分位文本。</summary>
        public string FormattedWanTotal { get; }

        /// <summary>按两位舍入后的合计对应的中文大写金额。</summary>
        public string ChineseTotal { get; }

        /// <summary>实际参与求和的可见数字单元格数量。</summary>
        public int NumericCellCount { get; }
    }

    /// <summary>
    /// 选区求和业务算法，不依赖 Excel/WPS COM。
    /// </summary>
    public static class SelectionSumService
    {
        public const int RoundDigits = 2;

        /// <summary>
        /// 对已由 Host 读取并分类的单元格求和。
        /// 隐藏行/列、非数值类型和无法安全转为 decimal 的单元格均跳过。
        /// </summary>
        public static SelectionSumResult Calculate(IReadOnlyList<CellData> cells)
        {
            if (cells == null)
            {
                throw new ArgumentNullException(nameof(cells));
            }

            decimal rawTotal = 0m;
            var numericCellCount = 0;

            for (var i = 0; i < cells.Count; i++)
            {
                var cell = cells[i];
                if (cell == null || cell.IsHidden || !cell.CellType.IsNumber())
                {
                    continue;
                }

                if (!cell.TryGetDecimal(out var value))
                {
                    continue;
                }

                rawTotal = checked(rawTotal + value);
                numericCellCount++;
            }

            var roundedTotal = RoundingService.Round(rawTotal, RoundDigits);
            var formattedTotal = roundedTotal.ToString("#,##0.00", CultureInfo.InvariantCulture);
            var formattedWanTotal = (roundedTotal / 10000m)
                .ToString("#,##0.00", CultureInfo.InvariantCulture);
            var chineseTotal = ChineseAmountService.ToChineseAmount(roundedTotal);

            return new SelectionSumResult(
                rawTotal,
                roundedTotal,
                formattedTotal,
                formattedWanTotal,
                chineseTotal,
                numericCellCount);
        }
    }
}
