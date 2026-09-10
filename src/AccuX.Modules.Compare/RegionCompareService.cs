using System;
using System.Collections.Generic;
using System.Globalization;
using AccuX.Core.Cells;
using AccuX.Core.Operations;

namespace AccuX.Modules.Compare
{
    /// <summary>
    /// 纯 CLR 的存在对比算法。它不读取当前选区，也不依赖 Excel/WPS。
    /// </summary>
    public sealed class RegionCompareService
    {
        public RegionCompareResult Compare(
            RangeReadResult first,
            RangeReadResult second,
            bool excludeHeaders)
        {
            if (first == null) throw new ArgumentNullException(nameof(first));
            if (second == null) throw new ArgumentNullException(nameof(second));

            var statistics = new RegionCompareStatistics { HeadersExcluded = excludeHeaders };
            var firstIndex = BuildIndex(first, excludeHeaders, statistics, true);
            var secondIndex = BuildIndex(second, excludeHeaders, statistics, false);

            var firstOnly = new List<RegionCompareItem>();
            var secondOnly = new List<RegionCompareItem>();
            var same = new List<RegionCompareItem>();

            foreach (var entry in firstIndex.Ordered)
            {
                if (secondIndex.ByKey.TryGetValue(entry.Key, out var other))
                {
                    same.Add(ToItem(entry, other, RegionCompareCategory.Same));
                }
                else
                {
                    firstOnly.Add(ToItem(entry, null, RegionCompareCategory.FirstOnly));
                }
            }

            foreach (var entry in secondIndex.Ordered)
            {
                if (!firstIndex.ByKey.ContainsKey(entry.Key))
                {
                    secondOnly.Add(ToItem(null, entry, RegionCompareCategory.SecondOnly));
                }
            }

            return new RegionCompareResult(firstOnly, secondOnly, same, statistics);
        }

        private static Index BuildIndex(
            RangeReadResult read,
            bool excludeHeaders,
            RegionCompareStatistics statistics,
            bool first)
        {
            var index = new Index();
            foreach (var cell in read.Cells)
            {
                if (cell == null || (excludeHeaders && cell.Row == 0))
                {
                    continue;
                }

                if (cell.IsHidden)
                {
                    if (first) statistics.FirstHiddenCount++; else statistics.SecondHiddenCount++;
                    continue;
                }

                if (cell.CellType == CellValueType.Blank || cell.Value == null
                    || (cell.Value is string empty && empty.Length == 0))
                {
                    if (first) statistics.FirstBlankCount++; else statistics.SecondBlankCount++;
                    continue;
                }

                if (!TryCreateKey(cell, out var key, out var display, out var type))
                {
                    if (first) statistics.FirstUnsupportedCount++; else statistics.SecondUnsupportedCount++;
                    continue;
                }

                if (!index.ByKey.TryGetValue(key, out var entry))
                {
                    entry = new Entry(key, display, type);
                    index.ByKey.Add(key, entry);
                    index.Ordered.Add(entry);
                }

                entry.Positions.Add(new RegionCompareCellPosition(cell.Row, cell.Column));
            }

            return index;
        }

        private static RegionCompareItem ToItem(Entry first, Entry second, RegionCompareCategory category)
        {
            var source = first ?? second;
            return new RegionCompareItem(
                source.Key,
                source.Display,
                source.Type,
                first == null ? (IReadOnlyList<RegionCompareCellPosition>)Array.Empty<RegionCompareCellPosition>() : first.Positions,
                second == null ? (IReadOnlyList<RegionCompareCellPosition>)Array.Empty<RegionCompareCellPosition>() : second.Positions,
                category);
        }

        private static bool TryCreateKey(
            CellData cell,
            out string key,
            out string display,
            out CellValueType type)
        {
            key = null;
            display = null;
            type = cell.CellType;

            var family = GetFamily(cell.CellType);
            if (family == null || cell.Value == null)
            {
                return false;
            }

            switch (family)
            {
                case "number":
                    if (!TryDecimal(cell.Value, out var number)) return false;
                    key = "number:" + number.ToString("G29", CultureInfo.InvariantCulture);
                    display = number.ToString(CultureInfo.CurrentCulture);
                    return true;
                case "text":
                    display = Convert.ToString(cell.Value, CultureInfo.InvariantCulture) ?? string.Empty;
                    key = "text:" + display;
                    return true;
                case "date":
                    if (cell.Value is DateTime date)
                    {
                        key = "date:" + date.Ticks.ToString(CultureInfo.InvariantCulture);
                        display = date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.CurrentCulture);
                        return true;
                    }

                    if (!TryDecimal(cell.Value, out var serial)) return false;
                    try
                    {
                        var serialDate = DateTime.FromOADate((double)serial);
                        key = "date:" + serialDate.Ticks.ToString(CultureInfo.InvariantCulture);
                        display = serialDate.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.CurrentCulture);
                        return true;
                    }
                    catch
                    {
                        return false;
                    }
                case "boolean":
                    if (!(cell.Value is bool boolean)) return false;
                    key = "boolean:" + (boolean ? "1" : "0");
                    display = boolean ? "TRUE" : "FALSE";
                    return true;
                case "error":
                    display = Convert.ToString(cell.Value, CultureInfo.InvariantCulture) ?? string.Empty;
                    key = "error:" + display;
                    return true;
                default:
                    return false;
            }
        }

        private static string GetFamily(CellValueType type)
        {
            switch (type)
            {
                case CellValueType.ConstantNumber:
                case CellValueType.FormulaNumber: return "number";
                case CellValueType.Text:
                case CellValueType.FormulaText: return "text";
                case CellValueType.Date:
                case CellValueType.FormulaDate: return "date";
                case CellValueType.Boolean:
                case CellValueType.FormulaBoolean: return "boolean";
                case CellValueType.Error:
                case CellValueType.FormulaError: return "error";
                default: return null;
            }
        }

        private static bool TryDecimal(object value, out decimal result)
        {
            switch (value)
            {
                case decimal d: result = d; return true;
                case double d: result = (decimal)d; return true;
                case float f: result = (decimal)f; return true;
                case int i: result = i; return true;
                case long l: result = l; return true;
                case short s: result = s; return true;
                case byte b: result = b; return true;
                default:
                    result = 0m;
                    return decimal.TryParse(value as string, NumberStyles.Any, CultureInfo.InvariantCulture, out result);
            }
        }

        private sealed class Index
        {
            public Dictionary<string, Entry> ByKey { get; } = new Dictionary<string, Entry>(StringComparer.Ordinal);
            public List<Entry> Ordered { get; } = new List<Entry>();
        }

        private sealed class Entry
        {
            public Entry(string key, string display, CellValueType type)
            {
                Key = key; Display = display; Type = type;
            }

            public string Key { get; }
            public string Display { get; }
            public CellValueType Type { get; }
            public List<RegionCompareCellPosition> Positions { get; } = new List<RegionCompareCellPosition>();
        }
    }
}
