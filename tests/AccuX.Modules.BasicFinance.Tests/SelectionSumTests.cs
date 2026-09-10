using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
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
            Assert.Equal("0.12", result.FormattedWanTotal);
            Assert.Equal("壹仟贰佰叁拾肆元伍角柒分", result.ChineseTotal);
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
            Assert.Equal("0.00", result.FormattedWanTotal);
            Assert.Equal("壹元零壹分", result.ChineseTotal);
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
            Assert.Equal("0.00", result.FormattedWanTotal);
            Assert.Equal("零元整", result.ChineseTotal);
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
        public void Execute_ReadsOnceShowsDialogAndSuppressesGlobalMessage()
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
            Assert.False(result.ShowMessage);
            Assert.Equal(1, prompt.DialogCalls);
            Assert.Equal("1,234.57", prompt.DialogInput.FormattedTotal);
            Assert.Equal("0.12", prompt.DialogInput.FormattedWanTotal);
            Assert.Equal("壹仟贰佰叁拾肆元伍角柒分", prompt.DialogInput.ChineseTotal);
            Assert.Equal(1, host.CaptureCalls);
            Assert.Equal(1, host.ReadCalls);
            Assert.Equal(1, prompt.ConfirmCalls);
            Assert.Equal(0, prompt.CopyCalls);
        }

        [Fact]
        public void Execute_CopyFailureDoesNotClaimSuccess()
        {
            var host = new RecordingHost(new List<CellData>
            {
                Cell(0, 0, CellValueType.ConstantNumber, 12m)
            }, 1, 1);
            var prompt = new RecordingPrompt
            {
                DialogResult = SelectionSumDialogResult.CopyAttempted(
                    SelectionSumCopyKind.Amount,
                    false)
            };
            var context = CreateContext(host);
            var command = new SelectionSumCommand(prompt);
            var definition = command.CreateDefinition("test.module");

            var result = definition.Handler(new CommandExecutionContext(definition, context));

            Assert.False(result.Success);
            Assert.False(result.ShowMessage);
            Assert.Equal("复制至剪切板失败，请重试。", result.Message);
            Assert.Equal(1, prompt.DialogCalls);
            Assert.Equal(0, prompt.CopyCalls);
        }

        [Fact]
        public void Execute_DialogCancellationSkipsCopyAndGlobalMessage()
        {
            var host = new RecordingHost(new List<CellData>
            {
                Cell(0, 0, CellValueType.ConstantNumber, 12m)
            }, 1, 1);
            var prompt = new RecordingPrompt
            {
                DialogResult = SelectionSumDialogResult.Cancelled()
            };
            var context = CreateContext(host);
            var command = new SelectionSumCommand(prompt);
            var definition = command.CreateDefinition("test.module");

            var result = definition.Handler(new CommandExecutionContext(definition, context));

            Assert.True(result.Success);
            Assert.False(result.ShowMessage);
            Assert.Equal(1, prompt.DialogCalls);
            Assert.Equal(0, prompt.CopyCalls);
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
            Assert.False(result.ShowMessage);
            Assert.Equal("0.00", prompt.DialogInput.FormattedTotal);
            Assert.Equal("0.00", prompt.DialogInput.FormattedWanTotal);
            Assert.Equal("零元整", prompt.DialogInput.ChineseTotal);
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

    public class SelectionSumWindowTests
    {
        [Theory]
        [InlineData("AmountButton", "1,234.57", SelectionSumCopyKind.Amount)]
        [InlineData("WanAmountButton", "0.12", SelectionSumCopyKind.WanAmount)]
        [InlineData("ChineseAmountButton", "壹仟贰佰叁拾肆元伍角柒分", SelectionSumCopyKind.ChineseAmount)]
        public void CopyButton_CopiesExpectedValueAndClosesWindow(
            string buttonName,
            string expectedText,
            SelectionSumCopyKind expectedKind)
        {
            Exception failure = null;
            string copiedText = null;
            SelectionSumDialogResult dialogResult = null;

            var thread = new Thread(() =>
            {
                try
                {
                    var window = new SelectionSumWindow(
                        CreateResult(),
                        IntPtr.Zero,
                        text =>
                        {
                            copiedText = text;
                            return true;
                        });

                    window.Loaded += (_, __) =>
                    {
                        Assert.True(((TextBox)window.FindName("AmountBox")).IsReadOnly);
                        Assert.True(((TextBox)window.FindName("WanAmountBox")).IsReadOnly);
                        Assert.True(((TextBox)window.FindName("ChineseAmountBox")).IsReadOnly);

                        var button = (Button)window.FindName(buttonName);
                        Assert.NotNull(button);
                        button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                    };

                    window.ShowDialog();
                    dialogResult = window.Result;
                }
                catch (Exception ex)
                {
                    failure = ex;
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            Assert.Null(failure);
            Assert.NotNull(dialogResult);
            Assert.False(dialogResult.WasCancelled);
            Assert.Equal(expectedKind, dialogResult.CopyKind);
            Assert.True(dialogResult.CopySucceeded);
            Assert.Equal(expectedText, copiedText);
        }

        [Fact]
        public void CopyFailure_StillClosesWindowWithoutMessageBox()
        {
            Exception failure = null;
            SelectionSumDialogResult dialogResult = null;

            var thread = new Thread(() =>
            {
                try
                {
                    var window = new SelectionSumWindow(
                        CreateResult(),
                        IntPtr.Zero,
                        text => false);

                    window.Loaded += (_, __) =>
                    {
                        var button = (Button)window.FindName("AmountButton");
                        button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                    };

                    window.ShowDialog();
                    dialogResult = window.Result;
                }
                catch (Exception ex)
                {
                    failure = ex;
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            Assert.Null(failure);
            Assert.NotNull(dialogResult);
            Assert.False(dialogResult.WasCancelled);
            Assert.Equal(SelectionSumCopyKind.Amount, dialogResult.CopyKind);
            Assert.False(dialogResult.CopySucceeded);
        }

        private static SelectionSumResult CreateResult()
        {
            return new SelectionSumResult(
                1234.567m,
                1234.57m,
                "1,234.57",
                "0.12",
                "壹仟贰佰叁拾肆元伍角柒分",
                2);
        }
    }

    internal sealed class RecordingPrompt : IUserPrompt
    {
        public bool ConfirmResult { get; set; } = true;

        public bool CopyResult { get; set; } = true;

        public SelectionSumDialogResult DialogResult { get; set; } =
            SelectionSumDialogResult.CopyAttempted(SelectionSumCopyKind.Amount, true);

        public int ConfirmCalls { get; private set; }

        public int CopyCalls { get; private set; }

        public int DialogCalls { get; private set; }

        public SelectionSumResult DialogInput { get; private set; }

        public RoundingOptions AskRoundingOptions(RangeTarget target)
        {
            return null;
        }

        public AmountConversionOptions AskAmountConversionOptions(RangeTarget target)
        {
            return null;
        }

        public SelectionSumDialogResult ShowSelectionSumDialog(SelectionSumResult result)
        {
            DialogCalls++;
            DialogInput = result;
            return DialogResult;
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
