using System;
using AccuX.Core.Operations;

namespace AccuX.Host
{
    /// <summary>
    /// <see cref="IHostContext"/> 的默认实现。
    /// </summary>
    public sealed class HostContext : IHostContext
    {
        public HostContext(HostKind hostKind, string hostVersion, IntPtr mainWindowHandle, object application)
        {
            HostKind = hostKind;
            HostVersion = hostVersion ?? string.Empty;
            MainWindowHandle = mainWindowHandle;
            Application = application;
        }

        public HostKind HostKind { get; }

        public string HostVersion { get; }

        public IntPtr MainWindowHandle { get; }

        /// <summary>仅限 AddIn / Host 边界内使用，不得传入业务 Module。</summary>
        public object Application { get; }
    }
}
