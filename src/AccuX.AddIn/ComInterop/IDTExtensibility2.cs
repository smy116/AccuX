using System;
using System.Runtime.InteropServices;

namespace AccuX.AddIn.ComInterop
{
    /// <summary>
    /// Office COM Add-in 生命周期接口（IDTExtensibility2）。
    /// <para>
    /// 该接口没有任何可用的 NuGet 包，因此在此显式声明，避免依赖本机 GAC 中版本不一致的 office.dll。
    /// GUID 与 Microsoft Add-in Designer 定义完全一致。
    /// </para>
    /// </summary>
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object application,
            [MarshalAs(UnmanagedType.I4)] ext_ConnectMode connectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object addInInst,
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(2)]
        void OnDisconnection(
            [MarshalAs(UnmanagedType.I4)] ext_DisconnectMode removeMode,
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(4)]
        void OnStartupComplete(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    /// <summary>加载方式。</summary>
    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3
    }

    /// <summary>卸载方式。</summary>
    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1
    }
}
