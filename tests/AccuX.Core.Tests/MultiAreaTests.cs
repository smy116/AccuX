using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using AccuX.Core.Operations;
using AccuX.Host;
using Xunit;
using Bounds = AccuX.Host.ExcelSelectionAreas.Bounds;

namespace AccuX.Core.Tests
{
    public class MultiAreaTests
    {
        private static RangeTarget Target() => new RangeTarget("book", "sheet", "Sheet1",
            new[] { new RangeArea("A1:B2", 2, 2), new RangeArea("F5", 1, 1) }, false);

        [Fact]
        public void Target_CopiesAreasAndPreservesIdentity()
        {
            var areas = new List<RangeArea> { new RangeArea("A1:B2", 2, 2), new RangeArea("F5", 1, 1) };
            var target = new RangeTarget("book", "sheet", "Sheet1", areas, false, "token");
            areas.Clear();
            Assert.Equal(5, target.CellCount);
            Assert.True(target.IsMultiArea);
            Assert.Equal(0, target.RowCount);
            Assert.Equal(2, target.Areas.Count);
            var single = target.GetAreaTarget(1);
            Assert.Equal("F5", single.Address);
            Assert.Equal("token", single.IdentityToken);
            Assert.Equal("book", single.WorkbookKey);
            Assert.False(single.IsMultiArea);
        }

        [Fact]
        public void SplitPlan_PreservesAreaCoordinatesValuesFormulasAndFormats()
        {
            var target = Target();
            var plan = new RangeWritePlan(target);
            plan.AddValue(0, 0, 12m, "0.00", 0);
            plan.AddFormula(0, 0, "=H1+1", 1);
            var result = Split(target, plan);
            Assert.Equal(12m, Assert.Single(result[0].Writes).Value);
            Assert.Equal("0.00", result[0].Writes[0].NumberFormat);
            Assert.Equal("=H1+1", Assert.Single(result[1].Writes).Formula);
            Assert.Equal(0, result[1].Writes[0].AreaIndex);
            Assert.Equal("F5", result[1].Target.Address);
        }

        [Theory]
        [InlineData(2, 0, 0)]
        [InlineData(-1, 0, 0)]
        [InlineData(1, 1, 0)]
        [InlineData(1, 0, 1)]
        public void SplitPlan_RejectsInvalidCoordinates(int area, int row, int column)
        {
            var target = Target();
            var plan = new RangeWritePlan(target);
            plan.AddValue(row, column, 1, areaIndex: area);
            var error = Assert.Throws<TargetInvocationException>(() => Split(target, plan));
            Assert.IsType<HostOperationException>(error.InnerException);
        }

        [Fact]
        public void SplitPlan_RejectsMismatchedTarget()
        {
            var error = Assert.Throws<TargetInvocationException>(() => Split(Target(), new RangeWritePlan(Target())));
            Assert.IsType<HostOperationException>(error.InnerException);
        }

        [Fact]
        public void Normalize_ClipsAndDeduplicatesOverlaps()
        {
            var result = ExcelSelectionAreas.Normalize(new[] {
                new Bounds(1, 2, 1, 2), new Bounds(2, 3, 2, 3),
                new Bounds(1, 2, 1, 2), new Bounds(20, 30, 20, 30)
            }, new Bounds(1, 10, 1, 10), 7);
            Assert.Equal(7, result.Sum(area => area.CellCount));
            var cells = Expand(result).ToArray();
            Assert.Equal(cells.Length, cells.Distinct().Count());
        }

        [Fact]
        public void Subtract_InteriorCutProducesFourDisjointRectangles()
        {
            var result = ExcelSelectionAreas.Subtract(new Bounds(1, 5, 1, 5), new Bounds(2, 4, 2, 4));
            Assert.Equal(4, result.Count);
            Assert.Equal(16, result.Sum(area => area.CellCount));
            Assert.Equal(16, Expand(result).Distinct().Count());
        }

        [Fact]
        public void Normalize_EnforcesCellLimitAndRejectsEmptyIntersection()
        {
            Assert.Throws<HostOperationException>(() => ExcelSelectionAreas.Normalize(
                new[] { new Bounds(1, 2, 1, 2) }, new Bounds(1, 10, 1, 10), 3));
            Assert.Throws<HostOperationException>(() => ExcelSelectionAreas.Normalize(
                new[] { new Bounds(20, 21, 20, 21) }, new Bounds(1, 10, 1, 10), 10));
            Assert.Throws<HostOperationException>(() => ExcelSelectionAreas.ValidateAreaCount(4097));
        }

        [Fact]
        public void Normalize_MatchesUnionAcrossOverlappingRectangles()
        {
            var random = new Random(42);
            for (var trial = 0; trial < 100; trial++)
            {
                var input = Enumerable.Range(0, 10).Select(_ => {
                    var row = random.Next(1, 9);
                    var col = random.Next(1, 9);
                    return new Bounds(row, row + random.Next(3), col, col + random.Next(3));
                }).ToArray();
                var expected = Expand(input).Distinct().OrderBy(value => value).ToArray();
                var normalized = ExcelSelectionAreas.Normalize(input, new Bounds(1, 10, 1, 10), 100);
                var actual = Expand(normalized).OrderBy(value => value).ToArray();
                Assert.Equal(expected, actual);
            }
        }

        private static RangeWritePlan[] Split(RangeTarget target, RangeWritePlan plan) =>
            (RangeWritePlan[])typeof(ExcelRangeOperationHost).GetMethod("SplitWritePlan",
                BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, new object[] { target, plan });

        private static IEnumerable<string> Expand(IEnumerable<Bounds> areas)
        {
            foreach (var area in areas)
                for (var row = area.FirstRow; row <= area.LastRow; row++)
                    for (var column = area.FirstColumn; column <= area.LastColumn; column++)
                        yield return row + ":" + column;
        }
    }
}
