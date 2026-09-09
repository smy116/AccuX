using System;
using System.Collections.Generic;
using AccuX.Core.Cells;
using AccuX.Core.Commands;
using AccuX.Core.Configuration;
using AccuX.Core.Logging;
using AccuX.Core.Modules;
using AccuX.Core.Operations;
using Xunit;

namespace AccuX.Core.Tests
{
    public class CommandDispatcherTests
    {
        private static IAccuXContext CreateContext(IRangeOperationHost host)
        {
            var pipeline = new RangeOperationPipeline(host, NullLogger.Instance);
            return new ModuleContext(
                new JsonConfigManager(System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + ".json")),
                NullLogger.Instance,
                host.Context,
                pipeline,
                "1.0.0");
        }

        [Fact]
        public void Execute_UnknownCommand_ReturnsFailure()
        {
            var dispatcher = new CommandDispatcher(CreateContext(new FakeHost()), NullLogger.Instance);
            var result = dispatcher.Execute("missing");

            Assert.False(result.Success);
        }

        [Fact]
        public void Execute_InvokesHandler()
        {
            var dispatcher = new CommandDispatcher(CreateContext(new FakeHost()), NullLogger.Instance);
            var invoked = false;

            dispatcher.Register(new CommandDefinition("test.cmd", "Test", "test", ctx =>
            {
                invoked = true;
                return CommandResult.Ok("done");
            }));

            var result = dispatcher.Execute("test.cmd");

            Assert.True(invoked);
            Assert.True(result.Success);
            Assert.Equal("done", result.Message);
        }

        [Fact]
        public void Execute_HandlerThrows_DoesNotPropagate()
        {
            var dispatcher = new CommandDispatcher(CreateContext(new FakeHost()), NullLogger.Instance);
            dispatcher.Register(new CommandDefinition("test.boom", "Boom", "test", ctx =>
            {
                throw new InvalidOperationException("kaboom");
            }));

            var result = dispatcher.Execute("test.boom");

            Assert.False(result.Success);
            Assert.Contains("kaboom", result.Message);
        }

        [Fact]
        public void Execute_HostOperationException_MessagePassedThrough()
        {
            var dispatcher = new CommandDispatcher(CreateContext(new FakeHost()), NullLogger.Instance);
            dispatcher.Register(new CommandDefinition("test.host", "Host", "test", ctx =>
            {
                throw new HostOperationException("请先选择一个数据区域。");
            }));

            var result = dispatcher.Execute("test.host");

            Assert.False(result.Success);
            Assert.Equal("请先选择一个数据区域。", result.Message);
        }
    }

    internal sealed class FakeHost : IRangeOperationHost
    {
        public IHostContext Context { get; } = new FakeHostContext();

        public RangeTarget CaptureTarget()
        {
            return TestData.Target();
        }

        public RangeReadResult Read(RangeTarget target)
        {
            return new RangeReadResult(target, new List<CellData>());
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
            return new FakeScope();
        }

        public string[,] ReadNumberFormats(RangeTarget target)
        {
            return new string[1, 1];
        }
    }

    internal sealed class FakeHostContext : IHostContext
    {
        public HostKind HostKind => HostKind.Excel;

        public string HostVersion => "16.0";

        public IntPtr MainWindowHandle => IntPtr.Zero;

        public object Application => null;
    }

    internal sealed class FakeScope : IHostStateScope
    {
        public void Dispose()
        {
        }
    }

    internal static class TestData
    {
        public static RangeTarget Target(int rows = 1, int columns = 1)
        {
            return new RangeTarget("book", "sheet", "Sheet1", "A1", rows, columns, (long)rows * columns, false, false);
        }

        public static CellData Cell(int row, int column, CellValueType type, object value)
        {
            return new CellData(row, column) { CellType = type, Value = value, Writable = true };
        }
    }
}
