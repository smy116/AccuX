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
            var prompt = new RecordingPrompt();

            var result = Execute(directoryHost, prompt);

            Assert.True(result.Success);
            Assert.Equal("生成完成！共生成3个表格的目录。", result.Message);
            Assert.Equal(1, directoryHost.GenerateCalls);
            Assert.False(directoryHost.LastReplaceExisting);
            Assert.Equal(0, prompt.ReplaceDirectoryConfirmCalls);
        }

        [Fact]
        public void Execute_PassesDirectoryTemplateToHost()
        {
            var directoryHost = new RecordingDirectoryHost { GeneratedCount = 1 };
            var prompt = new RecordingPrompt();

            var result = Execute(directoryHost, prompt);

            Assert.True(result.Success);
            var options = directoryHost.LastOptions;
            Assert.NotNull(options);
            Assert.Equal("目录", options.WorksheetName);
            Assert.Equal(new[] { "序号", "名称", "备注" }, options.Headers);
            Assert.Equal(3, options.ColumnWidths.Length);
            Assert.Equal(16, options.TitleFontSize);
            Assert.Equal(30d, options.TitleRowHeight);
            Assert.Equal("#143146", options.TitleBackgroundColor);
            Assert.Equal("#0563C1", options.HyperlinkColor);
        }

        [Fact]
        public void Execute_ExistingDirectoryConfirmed_RegeneratesWithReplacement()
        {
            var directoryHost = new RecordingDirectoryHost
            {
                GeneratedCount = 2,
                DirectoryExists = true
            };
            var prompt = new RecordingPrompt { ReplaceDirectoryResult = true };

            var result = Execute(directoryHost, prompt);

            Assert.True(result.Success);
            Assert.Equal("生成完成！共生成2个表格的目录。", result.Message);
            Assert.True(result.ShowMessage);
            Assert.Equal(1, prompt.ReplaceDirectoryConfirmCalls);
            Assert.Equal("目录", prompt.ReplaceDirectoryWorksheetName);
            Assert.Equal(1, directoryHost.GenerateCalls);
            Assert.True(directoryHost.LastReplaceExisting);
        }

        [Fact]
        public void Execute_ExistingDirectoryDeclined_ReturnsCancelledWithoutGenerating()
        {
            var directoryHost = new RecordingDirectoryHost
            {
                GeneratedCount = 2,
                DirectoryExists = true
            };
            var prompt = new RecordingPrompt { ReplaceDirectoryResult = false };

            var result = Execute(directoryHost, prompt);

            Assert.True(result.Success);
            Assert.False(result.ShowMessage);
            Assert.Equal(string.Empty, result.Message);
            Assert.Equal(1, prompt.ReplaceDirectoryConfirmCalls);
            Assert.Equal(0, directoryHost.GenerateCalls);
        }

        [Fact]
        public void Execute_WithoutWorkbookDirectoryHost_ReturnsFailure()
        {
            var context = CreateContext(null);
            var command = new DirectoryCommand(new RecordingPrompt());
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
                Failure = new HostOperationException("工作簿结构已保护，请先取消保护后重试。")
            };
            var prompt = new RecordingPrompt();

            var result = Execute(directoryHost, prompt);

            Assert.False(result.Success);
            Assert.Equal("工作簿结构已保护，请先取消保护后重试。", result.Message);
            Assert.Equal(1, directoryHost.GenerateCalls);
        }

        private static CommandResult Execute(RecordingDirectoryHost directoryHost, RecordingPrompt prompt)
        {
            var context = CreateContext(directoryHost);
            var command = new DirectoryCommand(prompt);
            var definition = command.CreateDefinition("test.module");

            return definition.Handler(new CommandExecutionContext(definition, context));
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

        public bool DirectoryExists { get; set; }

        public int GenerateCalls { get; private set; }

        public bool LastReplaceExisting { get; private set; }

        public DirectoryOptions LastOptions { get; private set; }

        public HostOperationException Failure { get; set; }

        public bool DirectoryWorksheetExists(string worksheetName)
        {
            return DirectoryExists;
        }

        public int GenerateDirectory(DirectoryOptions options, bool replaceExisting)
        {
            GenerateCalls++;
            LastOptions = options;
            LastReplaceExisting = replaceExisting;
            if (Failure != null)
            {
                throw Failure;
            }

            if (DirectoryExists && !replaceExisting)
            {
                throw new HostOperationException("工作簿中已存在“目录”工作表，请删除后重试。");
            }

            return GeneratedCount;
        }
    }
}
