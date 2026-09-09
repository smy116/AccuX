using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AccuX.Core.Commands;
using AccuX.Core.Configuration;
using AccuX.Core.Logging;
using AccuX.Core.Modules;
using AccuX.Core.Operations;
using AccuX.Modules.Mark;
using Xunit;
using static AccuX.Modules.Mark.Tests.MarkTestContextFactory;

namespace AccuX.Modules.Mark.Tests
{
    public class MarkModuleTests
    {
        [Fact]
        public void Initialize_RegistersFourCommandsInRequestedOrder()
        {
            var module = new MarkModule();
            module.Initialize(CreateContext(new RecordingMarkHost()));

            var commands = module.GetCommands().ToArray();

            Assert.Equal(4, commands.Length);
            Assert.Equal(
                new[]
                {
                    "accux.mark.green",
                    "accux.mark.red",
                    "accux.mark.yellow",
                    "accux.mark.blue"
                },
                commands.Select(command => command.Id).ToArray());
            Assert.Equal(new[] { "标绿", "标红", "标黄", "标蓝" }, commands.Select(command => command.DisplayName).ToArray());
            Assert.Contains("#18be6a", commands[0].Description, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#ed4015", commands[1].Description, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#fe9900", commands[2].Description, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#2db7f5", commands[3].Description, StringComparison.OrdinalIgnoreCase);
        }
    }

    public class MarkCommandTests
    {
        [Fact]
        public void Execute_CapturesOnce_UsesRequestedColor_AndIsSilentOnSuccess()
        {
            var rangeHost = new RecordingRangeHost();
            var markHost = new RecordingMarkHost { MarkedCount = 3 };
            var command = new MarkCommand("accux.mark.green", "标绿", "#18be6a");
            var definition = command.CreateDefinition("test.mark");

            var result = definition.Handler(new CommandExecutionContext(
                definition,
                CreateContext(markHost, rangeHost)));

            Assert.True(result.Success);
            Assert.False(result.ShowMessage);
            Assert.Equal(1, rangeHost.CaptureCalls);
            Assert.Equal(1, markHost.ApplyCalls);
            Assert.Equal("#18be6a", markHost.LastColor);
            Assert.Equal(rangeHost.Target, markHost.LastTarget);
        }

        [Fact]
        public void Execute_WhenNoVisibleCells_ShowsFriendlyMessage()
        {
            var markHost = new RecordingMarkHost { MarkedCount = 0 };
            var command = new MarkCommand("accux.mark.yellow", "标黄", "#fe9900");
            var definition = command.CreateDefinition("test.mark");

            var result = definition.Handler(new CommandExecutionContext(
                definition,
                CreateContext(markHost)));

            Assert.True(result.Success);
            Assert.True(result.ShowMessage);
            Assert.Contains("没有可见单元格", result.Message);
        }

        [Fact]
        public void Execute_WhenHostIsMissing_FailsBeforeCapturingSelection()
        {
            var rangeHost = new RecordingRangeHost();
            var command = new MarkCommand("accux.mark.blue", "标蓝", "#2db7f5");
            var definition = command.CreateDefinition("test.mark");

            var result = definition.Handler(new CommandExecutionContext(
                definition,
                CreateContext(null, rangeHost)));

            Assert.False(result.Success);
            Assert.Contains("不支持标记", result.Message);
            Assert.Equal(0, rangeHost.CaptureCalls);
        }

        [Fact]
        public void Execute_WhenCaptureFails_DoesNotWriteColor()
        {
            var rangeHost = new RecordingRangeHost
            {
                CaptureException = new HostOperationException("当前选区为空。")
            };
            var markHost = new RecordingMarkHost();
            var command = new MarkCommand("accux.mark.red", "标红", "#ed4015");
            var definition = command.CreateDefinition("test.mark");

            var result = definition.Handler(new CommandExecutionContext(
                definition,
                CreateContext(markHost, rangeHost)));

            Assert.False(result.Success);
            Assert.Equal("当前选区为空。", result.Message);
            Assert.Equal(0, markHost.ApplyCalls);
        }

        [Fact]
        public void Execute_WhenHostWriteFails_ReturnsFriendlyFailure()
        {
            var markHost = new RecordingMarkHost
            {
                ApplyException = new HostOperationException("目标工作表处于保护状态。")
            };
            var command = new MarkCommand("accux.mark.red", "标红", "#ed4015");
            var definition = command.CreateDefinition("test.mark");

            var result = definition.Handler(new CommandExecutionContext(
                definition,
                CreateContext(markHost)));

            Assert.False(result.Success);
            Assert.Equal("目标工作表处于保护状态。", result.Message);
        }
    }

    internal static class MarkTestContextFactory
    {
        public static IAccuXContext CreateContext(
            RecordingMarkHost markHost,
            RecordingRangeHost rangeHost = null)
        {
            rangeHost = rangeHost ?? new RecordingRangeHost();
            var configPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".json");
            var context = new ModuleContext(
                new JsonConfigManager(configPath),
                NullLogger.Instance,
                rangeHost.Context,
                new RangeOperationPipeline(rangeHost, NullLogger.Instance),
                "1.0.0",
                cellMarkHost: markHost);
            return context;
        }
    }

    internal sealed class RecordingMarkHost : ICellMarkHost
    {
        public long MarkedCount { get; set; }

        public Exception ApplyException { get; set; }

        public int ApplyCalls { get; private set; }

        public RangeTarget LastTarget { get; private set; }

        public string LastColor { get; private set; }

        public long ApplyVisibleBackgroundColor(RangeTarget target, string hexColor)
        {
            ApplyCalls++;
            LastTarget = target;
            LastColor = hexColor;
            if (ApplyException != null)
            {
                throw ApplyException;
            }

            return MarkedCount;
        }
    }

    internal sealed class RecordingRangeHost : IRangeOperationHost
    {
        public RecordingRangeHost()
        {
            Target = new RangeTarget("book", "sheet", "Sheet1", "A1:B2", 2, 2, 4, false, false);
            Context = new MarkTestHostContext();
        }

        public IHostContext Context { get; }

        public RangeTarget Target { get; }

        public int CaptureCalls { get; private set; }

        public Exception CaptureException { get; set; }

        public RangeTarget CaptureTarget()
        {
            CaptureCalls++;
            if (CaptureException != null)
            {
                throw CaptureException;
            }

            return Target;
        }

        public RangeReadResult Read(RangeTarget target)
        {
            return new RangeReadResult(target, Array.Empty<CellData>());
        }

        public WriteCheckResult ValidateWrite(RangeTarget target, RangeWritePlan writePlan)
        {
            return WriteCheckResult.Success();
        }

        public void Write(RangeTarget target, RangeWritePlan writePlan)
        {
        }

        public IHostStateScope BeginStateScope(HostStateOptions options)
        {
            return new MarkTestScope();
        }

        public string[,] ReadNumberFormats(RangeTarget target)
        {
            return new string[target.RowCount, target.ColumnCount];
        }
    }

    internal sealed class MarkTestHostContext : IHostContext
    {
        public HostKind HostKind => HostKind.Excel;

        public string HostVersion => "16.0";

        public IntPtr MainWindowHandle => IntPtr.Zero;

        public object Application => null;
    }

    internal sealed class MarkTestScope : IHostStateScope
    {
        public void Dispose()
        {
        }
    }
}
