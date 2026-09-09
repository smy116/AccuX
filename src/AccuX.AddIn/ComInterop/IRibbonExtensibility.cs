using System;
using System.Runtime.InteropServices;

namespace AccuX.AddIn.ComInterop
{
    /// <summary>
    /// Office Ribbon 扩展接口（IRibbonExtensibility）。
    /// GUID 与 Office 定义一致；同样显式声明以避免依赖 office.dll。
    /// </summary>
    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonId);
    }

    /// <summary>
    /// Ribbon 按钮回调接口（IRibbonControl）。
    /// </summary>
    [ComImport]
    [Guid("000C0395-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonControl
    {
        [DispId(1)]
        string Id { get; }

        [DispId(2)]
        object Context { get; }

        [DispId(3)]
        string Tag { get; }
    }
}
