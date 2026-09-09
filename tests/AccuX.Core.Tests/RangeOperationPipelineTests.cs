using System;
using System.Collections.Generic;
using AccuX.Core.Cells;
using AccuX.Core.Logging;
using AccuX.Core.Operations;
using Xunit;

namespace AccuX.Core.Tests
{
    /// <summary>
    /// 验证 Pipeline 中可脱离宿主测试的编排逻辑（规格 §26）。
    /// </summary>
    public class RangeOperationPipelineTests
    {
        [Fact]
        public void Execute_ReadsClassifiesAndWrites()
        {
            var host = new RecordingHost(new List<CellData>
            {
                TestData.Cell(0, 0, CellValueType.ConstantNumber, 1.005m),
                TestData.Cell(0, 1, CellValueType.Text, "abc")
            });

            var pipeline = new RangeOperationPipeline(host, NullLogger.Instance);
            var target = TestData.Target(1, 2);
            var context = new OperationContext("cmd", "mod", target);
            var transform = new PassthroughTransform();

            var result = pipeline.Execute(transform, context, new HostStateOptions());

            Assert.True(result.Success);
            Assert.Equal(2, result.Stats.TotalCells);
            Assert.Equal(1, result.Stats.ProcessedCells);
            Assert.Equal(1, result.Stats.SkippedCells);
            Assert.Single(host.WritePlan.Writes);
            Assert.True(host.Wrote);
        }

        [Fact]
        public void Execute_SkipsHiddenCellsBeforeTransforming()
        {
            var hiddenColumnCell = TestData.Cell(0, 1, CellValueType.ConstantNumber, 2m);
            hiddenColumnCell.IsHiddenColumn = true;

            var hiddenRowCell = TestData.Cell(1, 0, CellValueType.ConstantNumber, 3m);
            hiddenRowCell.IsHiddenRow = true;

            var hiddenRowAndColumnCell = TestData.Cell(1, 1, CellValueType.ConstantNumber, 4m);
            hiddenRowAndColumnCell.IsHiddenRow = true;
            hiddenRowAndColumnCell.IsHiddenColumn = true;

            var host = new RecordingHost(new List<CellData>
            {
                TestData.Cell(0, 0, CellValueType.ConstantNumber, 1m),
                hiddenColumnCell,
                hiddenRowCell,
                hiddenRowAndColumnCell
            });

            var transform = new AlwaysWriteTransform();
            var pipeline = new RangeOperationPipeline(host, NullLogger.Instance);
            var context = new OperationContext("cmd", "mod", TestData.Target(2, 2));

            var result = pipeline.Execute(transform, context, new HostStateOptions());

            Assert.True(result.Success);
            Assert.Equal(4, result.Stats.TotalCells);
            Assert.Equal(1, result.Stats.ProcessedCells);
            Assert.Equal(3, result.Stats.GetSkipCount(SkipReason.Hidden));
            Assert.True(result.Stats.ContainsHiddenRows);
            Assert.True(result.Stats.ContainsHiddenColumns);
            Assert.Equal(1, transform.Calls);
            Assert.Single(host.WritePlan.Writes);
            Assert.Equal(0, host.WritePlan.Writes[0].Row);
            Assert.Equal(0, host.WritePlan.Writes[0].Column);
            Assert.Contains("跳过隐藏行/列：3", result.Message);
        }

        [Fact]
        public void RangeReadResult_ReportsHiddenRowsAndColumns()
        {
            var hiddenRow = TestData.Cell(0, 0, CellValueType.ConstantNumber, 1m);
            hiddenRow.IsHiddenRow = true;

            var hiddenColumn = TestData.Cell(0, 1, CellValueType.ConstantNumber, 2m);
            hiddenColumn.IsHiddenColumn = true;

            var readResult = new RangeReadResult(
                TestData.Target(1, 2),
                new List<CellData> { hiddenRow, hiddenColumn });

            Assert.True(readResult.ContainsHiddenRows);
            Assert.True(readResult.ContainsHiddenColumns);
            Assert.Equal(2, readResult.HiddenCellCount);
        }

        [Fact]
        public void Execute_EmptyWritePlan_DoesNotWrite()
        {
            var host = new RecordingHost(new List<CellData>
            {
                TestData.Cell(0, 0, CellValueType.Text, "abc")
            });

            var pipeline = new RangeOperationPipeline(host, NullLogger.Instance);
            var context = new OperationContext("cmd", "mod", TestData.Target());

            var result = pipeline.Execute(new PassthroughTransform(), context, new HostStateOptions());

            Assert.True(result.Success);
            Assert.False(host.Wrote);
        }

        [Fact]
        public void Execute_ValidateWriteFails_DoesNotWrite()
        {
            var host = new RecordingHost(new List<CellData>
            {
                TestData.Cell(0, 0, CellValueType.ConstantNumber, 1m)
            })
            {
                CheckResult = WriteCheckResult.Failure("工作表受保护")
            };

            var pipeline = new RangeOperationPipeline(host, NullLogger.Instance);
            var context = new OperationContext("cmd", "mod", TestData.Target());

            var result = pipeline.Execute(new PassthroughTransform(), context, new HostStateOptions());

            Assert.False(result.Success);
            Assert.False(host.Wrote);
        }

        [Fact]
        public void Execute_TransformThrows_ReturnsFailureWithoutWriting()
        {
            var host = new RecordingHost(new List<CellData>
            {
                TestData.Cell(0, 0, CellValueType.ConstantNumber, 1m)
            });

            var pipeline = new RangeOperationPipeline(host, NullLogger.Instance);
            var context = new OperationContext("cmd", "mod", TestData.Target());

            var result = pipeline.Execute(new ThrowingTransform(), context, new HostStateOptions());

            Assert.False(result.Success);
            Assert.False(host.Wrote);
        }

        [Fact]
        public void Execute_DisposesStateScope_EvenWhenWriteThrows()
        {
            var host = new RecordingHost(new List<CellData>
            {
                TestData.Cell(0, 0, CellValueType.ConstantNumber, 1m)
            })
            {
                ThrowOnWrite = true
            };

            var pipeline = new RangeOperationPipeline(host, NullLogger.Instance);
            var context = new OperationContext("cmd", "mod", TestData.Target());

            var result = pipeline.Execute(new PassthroughTransform(), context, new HostStateOptions());

            Assert.False(result.Success);
            Assert.True(host.ScopeDisposed);
        }

        [Fact]
        public void Execute_FormulaTransform_RecordsFormulaWrite()
        {
            var host = new RecordingHost(new List<CellData>
            {
                new CellData(0, 0)
                {
                    CellType = CellValueType.FormulaNumber,
                    Value = 1m,
                    Formula = new FormulaInfo("=A1", FormulaKind.Normal, true),
                    Writable = true
                }
            });

            var pipeline = new RangeOperationPipeline(host, NullLogger.Instance);
            var context = new OperationContext("cmd", "mod", TestData.Target());

            var result = pipeline.Execute(new FormulaTransform(), context, new HostStateOptions());

            Assert.True(result.Success);
            Assert.Equal(1, result.Stats.FormulaWrites);
            Assert.True(host.WritePlan.Writes[0].IsFormulaWrite);
        }

        [Fact]
        public void Execute_CaptureTarget_DelegatesToHost()
        {
            var host = new RecordingHost(new List<CellData>());
            var pipeline = new RangeOperationPipeline(host, NullLogger.Instance);

            var target = pipeline.CaptureTarget();

            Assert.True(host.CaptureCalled);
            Assert.NotNull(target);
        }

        [Fact]
        public void Read_DelegatesToHostWithoutWriting()
        {
            var cells = new List<CellData>
            {
                TestData.Cell(0, 0, CellValueType.ConstantNumber, 12m)
            };
            var host = new RecordingHost(cells);
            var pipeline = new RangeOperationPipeline(host, NullLogger.Instance);
            var target = TestData.Target();

            var result = pipeline.Read(target);

            Assert.NotNull(result);
            Assert.Equal(1, host.ReadCalls);
            Assert.False(host.Wrote);
            Assert.Single(result.Cells);
        }

        [Fact]
        public void Execute_PropagatesNumberFormatToWritePlan()
        {
            var host = new RecordingHost(new List<CellData>
            {
                TestData.Cell(0, 0, CellValueType.ConstantNumber, 1m)
            });

            var pipeline = new RangeOperationPipeline(host, NullLogger.Instance);
            var context = new OperationContext("cmd", "mod", TestData.Target());

            var result = pipeline.Execute(new TextFormatTransform(), context, new HostStateOptions());

            Assert.True(result.Success);
            Assert.Equal("@", host.WritePlan.Writes[0].NumberFormat);
        }
    }

    internal sealed class TextFormatTransform : IRangeTransform
    {
        public string Id => "textformat";

        public bool MayWriteFormula => false;

        public CellTransformOutcome Transform(CellData cell, OperationContext context)
        {
            return CellTransformOutcome.WriteValue(cell, "壹元整", "@");
        }
    }

    internal sealed class RecordingHost : IRangeOperationHost
    {
        private readonly List<CellData> _cells;

        public RecordingHost(List<CellData> cells)
        {
            _cells = cells;
        }

        public IHostContext Context { get; } = new FakeHostContext();

        public WriteCheckResult CheckResult { get; set; } = WriteCheckResult.Success();

        public bool ThrowOnWrite { get; set; }

        public bool Wrote { get; private set; }

        public bool CaptureCalled { get; private set; }

        public int ReadCalls { get; private set; }

        public bool ScopeDisposed { get; private set; }

        public RangeWritePlan WritePlan { get; private set; }

        public RangeTarget CaptureTarget()
        {
            CaptureCalled = true;
            return TestData.Target();
        }

        public RangeReadResult Read(RangeTarget target)
        {
            ReadCalls++;
            return new RangeReadResult(target, _cells);
        }

        public WriteCheckResult ValidateWrite(RangeTarget target, RangeWritePlan writePlan)
        {
            return CheckResult;
        }

        public void Write(RangeTarget target, RangeWritePlan writePlan)
        {
            if (ThrowOnWrite)
            {
                throw new InvalidOperationException("write failed");
            }

            Wrote = true;
            WritePlan = writePlan;
        }

        public IHostStateScope BeginStateScope(HostStateOptions options)
        {
            return new RecordingScope(() => ScopeDisposed = true);
        }

        public string[,] ReadNumberFormats(RangeTarget target)
        {
            return new string[1, 1];
        }
    }

    internal sealed class RecordingScope : IHostStateScope
    {
        private readonly Action _onDispose;

        public RecordingScope(Action onDispose)
        {
            _onDispose = onDispose;
        }

        public void Dispose()
        {
            _onDispose?.Invoke();
        }
    }

    internal sealed class PassthroughTransform : IRangeTransform
    {
        public string Id => "passthrough";

        public bool MayWriteFormula => false;

        public CellTransformOutcome Transform(CellData cell, OperationContext context)
        {
            if (cell.CellType == CellValueType.ConstantNumber)
            {
                return CellTransformOutcome.WriteValue(cell, cell.Value);
            }

            return CellTransformOutcome.Skip(cell, SkipReason.Text);
        }
    }

    internal sealed class ThrowingTransform : IRangeTransform
    {
        public string Id => "throwing";

        public bool MayWriteFormula => false;

        public CellTransformOutcome Transform(CellData cell, OperationContext context)
        {
            throw new InvalidOperationException("transform failed");
        }
    }

    internal sealed class AlwaysWriteTransform : IRangeTransform
    {
        public string Id => "alwayswrite";

        public bool MayWriteFormula => false;

        public int Calls { get; private set; }

        public CellTransformOutcome Transform(CellData cell, OperationContext context)
        {
            Calls++;
            return CellTransformOutcome.WriteValue(cell, cell.Value);
        }
    }

    internal sealed class FormulaTransform : IRangeTransform
    {
        public string Id => "formula";

        public bool MayWriteFormula => true;

        public CellTransformOutcome Transform(CellData cell, OperationContext context)
        {
            return CellTransformOutcome.WriteFormula(cell, "=ROUND(A1,2)");
        }
    }
}
