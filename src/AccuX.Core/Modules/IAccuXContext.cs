using AccuX.Core.Configuration;
using AccuX.Core.Logging;
using AccuX.Core.Operations;

namespace AccuX.Core.Modules
{
    /// <summary>
    /// 模块可见的运行时上下文。
    /// <para>
    /// 模块通过它获得配置、日志、宿主上下文与统一的 Range 操作管线；
    /// 不得通过该上下文获取任何 COM 对象。
    /// </para>
    /// </summary>
    public interface IAccuXContext
    {
        IConfigManager Config { get; }

        ILogger Logger { get; }

        IHostContext Host { get; }

        RangeOperationPipeline Pipeline { get; }

        /// <summary>由 Host 实现的工作簿目录操作能力。</summary>
        IWorkbookDirectoryHost WorkbookDirectoryHost { get; }

        /// <summary>由 Host 实现的传统单元格批注操作能力。</summary>
        ICellCommentHost CellCommentHost { get; }

        /// <summary>由 Host 实现的当前选区可见单元格标记能力。</summary>
        ICellMarkHost CellMarkHost { get; }

        /// <summary>AccuX 版本号，用于日志与提示。</summary>
        string AccuXVersion { get; }
    }
}
