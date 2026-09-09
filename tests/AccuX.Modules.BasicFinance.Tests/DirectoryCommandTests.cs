using System;
using System.Collections.Generic;
using System.IO;
using AccuX.Core.Commands;
using AccuX.Core.Configuration;
using AccuX.Core.Logging;
using AccuX.Core.Modules;
using AccuX.Core.Operations;
using AccuX.Modules.BasicFinance.Directory;
using Xunit;

namespace AccuX.Modules.BasicFinance.Tests
{
    public class DirectoryCommandTests
    {
        [Fact]
        public void Execute_ReturnsExactCompletionMessageAndGeneratedCount()
        {
            var directoryHost = new RecordingDirectoryHost { GeneratedCount = 3 };
            var context = CreateContext(directoryHost);
            var command = new DirectoryCommand();
            var definition = command.CreateDefinition("test.module");

            var result = definition.Handler(new CommandExecutionContext(definition, context));

            Assert.True(result.Success);
            Assert.Equal("生成完成！共生成3个表格的目录。", result.Message);
            Assert.Equal(1, directoryHost.GenerateCalls);
        }

        [Fact]
        public void Execute_WithoutWorkbookDirectoryHost_ReturnsFailure()
        {
            var context = CreateContext(null);
            var command = new DirectoryCommand();
            var definition = command.CreateDefinition("test.module");

            var result = definition.Handler(new CommandExecutionContext(definition, context));

            Assert.False(result.Success);
            Assert.Equal("当前宿主不支持生成目录。", result.Message);
        }

        [Fact]
        public void Execute_HostFailurePassesFriendlyMessage()
        {
            var directoryHost = new RecordingDirectoryHost
            {
                Failure = new HostOperationException("工作簿中已存在“目录”工作表，请删除后重试。")
            };
            var context = CreateContext(directoryHost);
            var command = new DirectoryCommand();
            var definition = command.CreateDefinition("test.module");

            var result = definition.Handler(new CommandExecutionContext(definition, context));

            Assert.False(result.Success);
            Assert.Equal("工作簿中已存在“目录”工作表，请删除后重试。", result.Message);
            Assert.Equal(1, directoryHost.GenerateCalls);
        }

        private static IAccuXContext CreateContext(IWorkbookDirectoryHost directoryHost)
        {
            var rangeHost = new RecordingHost(new List<AccuX.Core.Operations.CellData>(), 1, 1);
            return new ModuleContext(
                new JsonConfigManager(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".json")),
                NullLogger.Instance,
                rangeHost.Context,
                new RangeOperationPipeline(rangeHost, NullLogger.Instance),
                "1.0.0",
                directoryHost);
        }
    }

    internal sealed class RecordingDirectoryHost : IWorkbookDirectoryHost
    {
        public int GeneratedCount { get; set; }

        public int GenerateCalls { get; private set; }

        public HostOperationException Failure { get; set; }

        public int GenerateDirectory()
        {
            GenerateCalls++;
            if (Failure != null)
            {
                throw Failure;
            }

            return GeneratedCount;
        }
    }
}
