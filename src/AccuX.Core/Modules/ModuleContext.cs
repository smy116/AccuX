using System;
using AccuX.Core.Configuration;
using AccuX.Core.Logging;
using AccuX.Core.Operations;

namespace AccuX.Core.Modules
{
    /// <summary>
    /// <see cref="IAccuXContext"/> 的默认实现，由 AccuX.AddIn 在组合根创建。
    /// </summary>
    public sealed class ModuleContext : IAccuXContext
    {
        public ModuleContext(
            IConfigManager config,
            ILogger logger,
            IHostContext host,
            RangeOperationPipeline pipeline,
            string accuXVersion,
            IWorkbookDirectoryHost workbookDirectoryHost = null,
            ICellCommentHost cellCommentHost = null)
        {
            Config = config ?? throw new ArgumentNullException(nameof(config));
            Logger = logger ?? NullLogger.Instance;
            Host = host;
            Pipeline = pipeline ?? throw new ArgumentNullException(nameof(pipeline));
            WorkbookDirectoryHost = workbookDirectoryHost;
            CellCommentHost = cellCommentHost;
            AccuXVersion = accuXVersion ?? string.Empty;
        }

        public IConfigManager Config { get; }

        public ILogger Logger { get; }

        public IHostContext Host { get; }

        public RangeOperationPipeline Pipeline { get; }

        public IWorkbookDirectoryHost WorkbookDirectoryHost { get; }

        public ICellCommentHost CellCommentHost { get; }

        public string AccuXVersion { get; }
    }
}
