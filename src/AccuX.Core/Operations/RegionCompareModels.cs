using System;
using System.Collections.Generic;
using AccuX.Core.Cells;

namespace AccuX.Core.Operations
{
    /// <summary>区域对比结果中的一个单元格坐标（相对于所属区域，0 基）。</summary>
    public sealed class RegionCompareCellPosition
    {
        public RegionCompareCellPosition(int row, int column)
        {
            Row = row;
            Column = column;
        }

        public int Row { get; }

        public int Column { get; }
    }

    public enum RegionCompareCategory
    {
        FirstOnly,
        SecondOnly,
        Same
    }

    /// <summary>一个去重后的存在对比项，保留两侧全部出现位置。</summary>
    public sealed class RegionCompareItem
    {
        public RegionCompareItem(
            string key,
            string displayValue,
            CellValueType cellType,
            IReadOnlyList<RegionCompareCellPosition> firstPositions,
            IReadOnlyList<RegionCompareCellPosition> secondPositions,
            RegionCompareCategory category)
        {
            Key = key ?? string.Empty;
            DisplayValue = displayValue ?? string.Empty;
            CellType = cellType;
            FirstPositions = firstPositions ?? Array.Empty<RegionCompareCellPosition>();
            SecondPositions = secondPositions ?? Array.Empty<RegionCompareCellPosition>();
            Category = category;
        }

        public string Key { get; }

        public string DisplayValue { get; }

        public CellValueType CellType { get; }

        public IReadOnlyList<RegionCompareCellPosition> FirstPositions { get; }

        public IReadOnlyList<RegionCompareCellPosition> SecondPositions { get; }

        public int FirstCount { get { return FirstPositions.Count; } }

        public int SecondCount { get { return SecondPositions.Count; } }

        public RegionCompareCategory Category { get; }
    }

    public sealed class RegionCompareStatistics
    {
        public int FirstHiddenCount { get; set; }
        public int SecondHiddenCount { get; set; }
        public int FirstBlankCount { get; set; }
        public int SecondBlankCount { get; set; }
        public int FirstUnsupportedCount { get; set; }
        public int SecondUnsupportedCount { get; set; }
        public bool HeadersExcluded { get; set; }
    }

    public sealed class RegionCompareResult
    {
        public RegionCompareResult(
            IReadOnlyList<RegionCompareItem> firstOnly,
            IReadOnlyList<RegionCompareItem> secondOnly,
            IReadOnlyList<RegionCompareItem> same,
            RegionCompareStatistics statistics)
        {
            FirstOnly = firstOnly ?? Array.Empty<RegionCompareItem>();
            SecondOnly = secondOnly ?? Array.Empty<RegionCompareItem>();
            Same = same ?? Array.Empty<RegionCompareItem>();
            Statistics = statistics ?? new RegionCompareStatistics();
        }

        public IReadOnlyList<RegionCompareItem> FirstOnly { get; }
        public IReadOnlyList<RegionCompareItem> SecondOnly { get; }
        public IReadOnlyList<RegionCompareItem> Same { get; }
        public RegionCompareStatistics Statistics { get; }
    }

    /// <summary>供 Host 导出到新工作簿的纯 CLR 行数据。</summary>
    public sealed class RegionCompareExportRow
    {
        public string Category { get; set; }
        public string Value { get; set; }
        public string CellType { get; set; }
        public int FirstCount { get; set; }
        public int SecondCount { get; set; }
        public string FirstLocations { get; set; }
        public string SecondLocations { get; set; }
        public string FirstWorkbook { get; set; }
        public string FirstWorksheet { get; set; }
        public string FirstRange { get; set; }
        public string SecondWorkbook { get; set; }
        public string SecondWorksheet { get; set; }
        public string SecondRange { get; set; }
    }

    public sealed class RegionCompareExportData
    {
        public RegionCompareExportData(IReadOnlyList<RegionCompareExportRow> rows, string summary)
        {
            Rows = rows ?? Array.Empty<RegionCompareExportRow>();
            Summary = summary ?? string.Empty;
        }

        public IReadOnlyList<RegionCompareExportRow> Rows { get; }

        public string Summary { get; }
    }
}
