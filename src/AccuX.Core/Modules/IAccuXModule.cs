using System.Collections.Generic;

namespace AccuX.Core.Modules
{
    /// <summary>
    /// 业务模块契约（规格 §6）。
    /// V1 只保留最小生命周期与 Command 注册能力；不包含 Name/Version/Order/GetRibbonGroup 等无真实运行时需求的信息。
    /// </summary>
    public interface IAccuXModule
    {
        /// <summary>模块唯一标识，例如 AccuX.Modules.BasicFinance。</summary>
        string Id { get; }

        /// <summary>初始化模块。异常由 ModuleRegistry 捕获隔离，不得导致整个 Add-in 崩溃。</summary>
        void Initialize(IAccuXContext context);

        /// <summary>返回本模块注册的命令定义。</summary>
        IEnumerable<Commands.CommandDefinition> GetCommands();

        /// <summary>释放模块资源。</summary>
        void Shutdown();
    }
}
