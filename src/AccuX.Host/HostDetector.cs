using System;
using AccuX.Core.Operations;
using Excel = Microsoft.Office.Interop.Excel;

namespace AccuX.Host
{
    /// <summary>
    /// 宿主识别（规格 §5.2）。
    /// Excel 与 WPS 表格都暴露 Excel 兼容的 COM 对象模型，因此通过名称/安装路径/版本特征区分。
    /// </summary>
    public static class HostDetector
    {
        /// <summary>
        /// 根据 COM Application 对象识别宿主。
        /// </summary>
        public static IHostContext Detect(object application)
        {
            if (application == null)
            {
                return new HostContext(HostKind.Unknown, string.Empty, IntPtr.Zero, null);
            }

            // Office 通过 IDispatch 传入的 RCW 是 __ComObject，反射取不到属性；
            // 必须按 Excel 兼容接口做 QueryInterface 转换。
            var excel = application as Excel.Application;
            if (excel == null)
            {
                return new HostContext(HostKind.Unknown, string.Empty, IntPtr.Zero, application);
            }

            var name = TryGet(() => excel.Name);
            var path = TryGet(() => excel.Path);
            var version = TryGet(() => excel.Version);
            var handle = TryGetHandle(excel);

            var kind = DetectKind(name, path);

            return new HostContext(kind, version, handle, excel);
        }

        /// <summary>
        /// 按宿主名称与安装路径判定宿主类型（纯逻辑，可单元测试）。
        /// </summary>
        public static HostKind DetectKind(string name, string path)
        {
            var probe = ((name ?? string.Empty) + " " + (path ?? string.Empty)).ToLowerInvariant();

            if (probe.IndexOf("wps", StringComparison.Ordinal) >= 0 ||
                probe.IndexOf("kingsoft", StringComparison.Ordinal) >= 0)
            {
                return HostKind.Wps;
            }

            if (probe.IndexOf("excel", StringComparison.Ordinal) >= 0)
            {
                return HostKind.Excel;
            }

            // 无法识别时按 Excel 兼容路径处理：AccuX 只通过 Excel 兼容 COM 模型访问宿主。
            return HostKind.Unknown;
        }

        private static string TryGet(Func<string> getter)
        {
            try
            {
                return getter() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static IntPtr TryGetHandle(Excel.Application excel)
        {
            try
            {
                return new IntPtr(excel.Hwnd);
            }
            catch
            {
                return IntPtr.Zero;
            }
        }
    }
}
