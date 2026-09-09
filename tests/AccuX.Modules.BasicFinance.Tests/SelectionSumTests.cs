using System;
using System.Collections.Generic;
using System.IO;
using AccuX.Core.Cells;
using AccuX.Core.Commands;
using AccuX.Core.Configuration;
using AccuX.Core.Logging;
using AccuX.Core.Modules;
using AccuX.Core.Operations;
using AccuX.Modules.BasicFinance.AmountConversion;
using AccuX.Modules.BasicFinance.Common;
using AccuX.Modules.BasicFinance.Rounding;
using AccuX.Modules.BasicFinance.SelectionSum;
using Xunit;

namespace AccuX.Modules.BasicFinance.Tests
{
    public class SelectionSumServiceTests
    {
        [Fact]
        public void Calculate_IncludesFormulaNumbersAndSkipsHiddenAndNonNumbers()
        {
            var result = SelectionSumService.Calculate(new List<CellData>
            {
                Cell(0, 0, CellValueType.ConstantNumber, 1000.125m),
                Cell(0, 1, CellValueType.FormulaNumber, 234.442m),
                Cell(1, 0, CellValueType.ConstantNumber, 999m, hiddenRow: true),
                Cell(1, 1, CellValueType.FormulaNumber, 888m, hiddenColumn: true),
                Cell(2, 0, CellValueType.Text, "123"),
                Cell(2, 1, CellValueType.Date, 456m),
                Cell(3, 0, CellValueType.Boolean, true),
                Cell(3, 1, CellValueType.Blank, null),
                Cell(4, 0, CellValueType.Error, "#N/A")
            });

            Assert.Equal(1234.567m, result.RawTotal);
            Assert.Equal(1234.57m, result.RoundedTotal);
            Assert.Equal("1,234.57", result.FormattedTotal);
            Assert.Equal(2, result.NumericCellCount);
        }

        [Fact]
        public void Calculate_RoundsAfterSumming()
        {
            var result = SelectionSumService.Calculate(new List<CellData>
            {
                Cell(0, 0, CellValueType.ConstantNumber, 1.004m),
                Cell(0, 1, CellValueType.ConstantNumber, 0.001m)
            });

            Assert.Equal(1.005m, result.RawTotal);
            Assert.Equal(1.01m, result.RoundedTotal);
            Assert.Equal("1.01", result.FormattedTotal);
        }

        [Fact]
        public void Calculate_UsesAwayFromZeroForNegativeMidpoint()
        {
            var result = SelectionSumService.Calculate(new List<CellData>
            {
                Cell(0, 0, CellValueType.ConstantNumber, -1.005m)
            });

            Assert.Equal(-1.01m, result.RoundedTotal);
            Assert.Equal("-1.01", result.FormattedTotal);
        }

        [Fact]
        public void Calculate_EmptyOrNonNumericRangeReturnsZero()
        {
            var result = SelectionSumService.Calculate(new List<CellData>
            {
                Cell(0, 0, CellValueType.Text, "not a number"),
                Cell(0, 1, CellValueType.Blank, null)
            });

            Assert.Equal(0m, result.RawTotal);
            Assert.Equal(0m, result.RoundedTotal);
            Assert.Equal("0.00", result.FormattedTotal);
            Assert.Equal(0, result.NumericCellCount);
        }

        [Fact]
        public void Calculate_FormatsThousandsAndTrailingZeros()
        {
            var result = SelectionSumService.Calculate(new List<CellData>
            {
                Cell(0, 0, CellValueType.ConstantNumber, 1234567.8m)
            });

            Assert.Equal("1,234,567.80", result.FormattedTotal);
        }

        private static CellData Cell(
            int row,
            int column,
            CellValueType type,
            object value,
            bool hiddenRow = false,
            bool hiddenColumn = false)
        {
            return new CellData(row, column)
            {
                CellType = type,
                Value = value,
                IsHiddenRow = hiddenRow,
                IsHiddenColumn = hiddenColumn,
                Writable = true
            };
        }
    }

    public class SelectionSumCommandTests
    {
        [Fact]
        public void Execute_ReadsOnceCopiesFormattedTotalAndReturnsExactMessage()
        {
            var host = new RecordingHost(new List<CellData>
            {
                Cell(0, 0, CellValueType.ConstantNumber, 1000.125m),
                Cell(0, 1, CellValueType.FormulaNumber, 234.442m),
                Cell(1, 0, CellValueType.ConstantNumber, 900m, hiddenRow: true)
            }, 2, 2);
            var prompt = new RecordingPrompt();
            var context = CreateContext(host);
            var command = new SelectionSumCommand(prompt);
            var definition = command.CreateDefinition("test.module");

            var result = definition.Handler(new CommandExecutionContext(definition, context));

            Assert.True(result.Success);
            Assert.Equal("选定区域合计1,234.57，已复制至剪切板。", result.Message);
            Assert.Equal("1,234.57", prompt.CopiedText);
            Assert.Equal(1, host.CaptureCalls);
            Assert.Equal(1, host.ReadCalls);
            Assert.Equal(1, prompt.ConfirmCalls);
            Assert.Equal(1, prompt.CopyCalls);
        }

        [Fact]
        public void Execute_CopyFailureDoesNotClaimSuccess()
        {
            var host = new RecordingHost(new List<CellData>
            {
                Cell(0, 0, CellValueType.ConstantNumber, 12m)
            }, 1, 1);
            var prompt = new RecordingPrompt { CopyResult = false };
            var context = CreateContext(host);
            var command = new SelectionSumCommand(prompt);
            var definition = command.CreateDefinition("test.module");

            var result = definition.Handler(new CommandExecutionContext(definition, context));

            Assert.False(result.Success);
            Assert.Equal("合计已计算，但复制至剪切板失败，请重试。", result.Message);
            Assert.DoesNotContain("已复制", result.Message);
            Assert.Equal("12.00", prompt.CopiedText);
        }

        [Fact]
        public void Execute_ConfirmCancellationSkipsReadAndCopy()
        {
            var host = new RecordingHost(new List<CellData>
            {
                Cell(0, 0, CellValueType.ConstantNumber, 12m)
            }, 1, 1);
            var prompt = new RecordingPrompt { ConfirmResult = false };
            var context = CreateContext(host);
            var command = new SelectionSumCommand(prompt);
            var definition = command.CreateDefinition("test.module");

            var result = definition.Handler(new CommandExecutionContext(definition, context));

            Assert.True(result.Success);
            Assert.False(result.ShowMessage);
            Assert.Equal(1, host.CaptureCalls);
            Assert.Equal(0, host.ReadCalls);
            Assert.Equal(0, prompt.CopyCalls);
        }

        [Fact]
        public void Execute_NoNumericCellsCopiesZero()
        {
            var host = new RecordingHost(new List<CellData>
            {
                Cell(0, 0, CellValueType.Text, "123"),
                Cell(0, 1, CellValueType.Blank, null)
            }, 1, 2);
            var prompt = new RecordingPrompt();
            var context = CreateContext(host);
            var command = new SelectionSumCommand(prompt);
            var definition = command.CreateDefinition("test.module");

            var result = definition.Handler(new CommandExecutionContext(definition, context));

            Assert.True(result.Success);
            Assert.Equal("0.00", prompt.CopiedText);
            Assert.Equal("选定区域合计0.00，已复制至剪切板。", result.Message);
        }

        [Fact]
        public void Execute_HostReadFailureReturnsFriendlyFailure()
        {
            var host = new RecordingHost(new List<CellData>(), 1, 1)
            {
                ReadException = new HostOperationException("原选区已失效，请重新执行。")
            };
            var prompt = new RecordingPrompt();
            var context = CreateContext(host);
            var command = new SelectionSumCommand(prompt);
            var definition = command.CreateDefinition("test.module");

            var result = definition.Handler(new CommandExecutionContext(definition, context));

            Assert.False(result.Success);
            Assert.Equal("原选区已失效，请重新执行。", result.Message);
            Assert.Equal(0, prompt.CopyCalls);
        }

        private static IAccuXContext CreateContext(RecordingHost host)
        {
            return new ModuleContext(
                new JsonConfigManager(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".json")),
                NullLogger.Instance,
                host.Context,
                new RangeOperationPipeline(host, NullLogger.Instance),
                "1.0.0");
        }

        private static CellData Cell(
            int row,
            int column,
            CellValueType type,
            object value,
            bool hiddenRow = false,
            bool hiddenColumn = false)
        {
            return new CellData(row, column)
            {
                CellType = type,
                Value = value,
                IsHiddenRow = hiddenRow,
                IsHiddenColumn = hiddenColumn,
                Writable = true
            };
        }
    }

    internal sealed class RecordingPrompt : IUserPrompt
    {
        public bool ConfirmResult { get; set; } = true;

        public bool CopyResult { get; set; } = true;

        public int ConfirmCalls { get; private set; }

        public int CopyCalls { get; private set; }

        public string CopiedText { get; private set; }

        public RoundingOptions AskRoundingOptions(RangeTarget target)
        {
            return null;
        }

        public AmountConversionOptions AskAmountConversionOptions(RangeTarget target)
        {
            return null;
        }

        public bool ConfirmLargeSelection(RangeTarget target)
        {
            ConfirmCalls++;
            return ConfirmResult;
        }

        public void ShowMessage(string message)
        {
        }

        public void ShowError(string message)
        {
        }

        public bool TryCopyToClipboard(string text)
        {
            CopyCalls++;
            CopiedText = text;
            return CopyResult;
        }
    }

    internal sealed class RecordingHost : IRangeOperationHost
    {
        private readonly IReadOnlyList<CellData> _cells;
        private readonly RangeTarget _target;

        public RecordingHost(IReadOnlyList<CellData> cells, int rows, int columns)
        {
            _cells = cells;
            _target = new RangeTarget(
                "book",
                "sheet",
                "Sheet1",
                "A1",
                rows,
                columns,
                (long)rows * columns,
                false,
                false);
        }

        public IHostContext Context { get; } = new FakeHostContext();

        public int CaptureCalls { get; private set; }

        public int ReadCalls { get; private set; }

        public Exception ReadException { get; set; }

        public RangeTarget CaptureTarget()
        {
            CaptureCalls++;
            return _target;
        }

        public RangeReadResult Read(RangeTarget target)
        {
            ReadCalls++;
            if (ReadException != null)
            {
                throw ReadException;
            }

            return new RangeReadResult(target, _cells);
        }

        public WriteCheckResult ValidateWrite(RangeTarget target, RangeWritePlan writePlan)
        {
            return WriteCheckResult.Success();
        }

        public void Write(RangeTarget target, RangeWritePlan writePlan)
        {
            throw new InvalidOperationException("选区求和不应写回工作表。");
        }

        public IHostStateScope BeginStateScope(HostStateOptions options)
        {
            throw new InvalidOperationException("选区求和不应创建宿主状态作用域。");
        }

        public string[,] ReadNumberFormats(RangeTarget target)
        {
            return new string[target.RowCount, target.ColumnCount];
        }
    }

    internal sealed class FakeHostContext : IHostContext
    {
        public HostKind HostKind => HostKind.Excel;

        public string HostVersion => "16.0";

        public IntPtr MainWindowHandle => IntPtr.Zero;

        public object Application => null;
    }
}
