using System;
using System.Collections.Generic;
using System.Linq;
using AccuX.Core.Cells;
using AccuX.Core.Operations;
using AccuX.Modules.Compare;
using Xunit;

namespace AccuX.Modules.Compare.Tests
{
    public sealed class RegionCompareServiceTests
    {
        [Fact]
        public void Compare_ReturnsExistenceCategoriesAndDuplicateCounts()
        {
            var first = Read(
                Cell(0, 0, CellValueType.Text, "D"), Cell(1, 0, CellValueType.Text, "B"),
                Cell(2, 0, CellValueType.Text, "C"), Cell(3, 0, CellValueType.Text, "A"),
                Cell(4, 0, CellValueType.Text, "C"));
            var second = Read(
                Cell(0, 0, CellValueType.Text, "Z"), Cell(1, 0, CellValueType.Text, "A"),
                Cell(2, 0, CellValueType.Text, "C"), Cell(3, 0, CellValueType.Text, "F"));

            var result = new RegionCompareService().Compare(first, second, false);

            Assert.Equal(new[] { "D", "B" }, result.FirstOnly.Select(x => x.DisplayValue).ToArray());
            Assert.Equal(new[] { "Z", "F" }, result.SecondOnly.Select(x => x.DisplayValue).ToArray());
            Assert.Equal(new[] { "C", "A" }, result.Same.Select(x => x.DisplayValue).ToArray());
            Assert.Equal(2, result.Same.Single(x => x.DisplayValue == "C").FirstCount);
        }

        [Fact]
        public void Compare_DistinguishesTypesCaseAndWhitespace()
        {
            var first = Read(
                Cell(0, 0, CellValueType.ConstantNumber, 1m),
                Cell(1, 0, CellValueType.Text, "Abc"),
                Cell(2, 0, CellValueType.Text, "x "));
            var second = Read(
                Cell(0, 0, CellValueType.Text, "1"),
                Cell(1, 0, CellValueType.Text, "abc"),
                Cell(2, 0, CellValueType.Text, "x"));

            var result = new RegionCompareService().Compare(first, second, false);

            Assert.Equal(3, result.FirstOnly.Count);
            Assert.Equal(3, result.SecondOnly.Count);
            Assert.Empty(result.Same);
        }

        [Fact]
        public void Compare_FormulaResultMatchesConstantOfSameFamily()
        {
            var formula = Cell(0, 0, CellValueType.FormulaNumber, 12.5m);
            formula.Formula = new FormulaInfo("=A1+B1", FormulaKind.Normal, true);
            var result = new RegionCompareService().Compare(
                Read(formula),
                Read(Cell(0, 0, CellValueType.ConstantNumber, 12.50m)),
                false);

            Assert.Single(result.Same);
            Assert.Empty(result.FirstOnly);
            Assert.Empty(result.SecondOnly);
        }

        [Fact]
        public void Compare_SkipsHiddenBlankUnsupportedAndOptionalHeader()
        {
            var hidden = Cell(2, 0, CellValueType.Text, "hidden"); hidden.IsHiddenRow = true;
            var unsupported = Cell(3, 0, CellValueType.Unsupported, new object());
            var first = Read(Cell(0, 0, CellValueType.Text, "header"), Cell(1, 0, CellValueType.Text, "kept"), hidden, unsupported, Cell(4, 0, CellValueType.Blank, null));
            var second = Read(Cell(0, 0, CellValueType.Text, "other"), Cell(1, 0, CellValueType.Text, "kept"));

            var result = new RegionCompareService().Compare(first, second, true);

            Assert.Single(result.Same);
            Assert.Equal(1, result.Statistics.FirstHiddenCount);
            Assert.Equal(1, result.Statistics.FirstUnsupportedCount);
            Assert.Equal(1, result.Statistics.FirstBlankCount);
            Assert.True(result.Statistics.HeadersExcluded);
        }

        [Fact]
        public void Compare_UsesDateAndBooleanAndErrorFamilies()
        {
            var first = Read(
                Cell(0, 0, CellValueType.Date, new DateTime(2025, 1, 2)),
                Cell(1, 0, CellValueType.Boolean, true),
                Cell(2, 0, CellValueType.Error, -2146826281));
            var second = Read(
                Cell(0, 0, CellValueType.FormulaDate, new DateTime(2025, 1, 2)),
                Cell(1, 0, CellValueType.FormulaBoolean, true),
                Cell(2, 0, CellValueType.FormulaError, -2146826281));

            var result = new RegionCompareService().Compare(first, second, false);

            Assert.Equal(3, result.Same.Count);
        }

        [Fact]
        public void Compare_NormalizesNumericDateSerials()
        {
            var date = new DateTime(2025, 1, 2);
            var serial = (decimal)date.ToOADate();
            var result = new RegionCompareService().Compare(
                Read(Cell(0, 0, CellValueType.Date, serial)),
                Read(Cell(0, 0, CellValueType.FormulaDate, date)),
                false);

            Assert.Single(result.Same);
        }

        private static RangeReadResult Read(params CellData[] cells)
        {
            var target = new RangeTarget("book", Guid.NewGuid().ToString(), "Sheet1", "A1:A10", 10, 1, 10, false, false);
            return new RangeReadResult(target, cells);
        }

        private static CellData Cell(int row, int column, CellValueType type, object value)
        {
            return new CellData(row, column) { CellType = type, Value = value, Writable = true };
        }
    }
}
