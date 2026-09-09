using System;

namespace AccuX.Core.Operations
{
    /// <summary>
    /// 宿主类型。
    /// </summary>
    public enum HostKind
    {
        Unknown = 0,
        Excel = 1,
        Wps = 2
    }

    /// <summary>
    /// 宿主上下文（规格 §8）。仅用于初始化、WPF Owner 与兼容性判断。
    /// <see cref="Application"/> 只供 AddIn / Host 边界内使用，不得传入业务 Module。
    /// </summary>
    public interface IHostContext
    {
        HostKind HostKind { get; }

        string HostVersion { get; }

        IntPtr MainWindowHandle { get; }

        object Application { get; }
    }
}
