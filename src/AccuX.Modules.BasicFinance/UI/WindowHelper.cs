using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace AccuX.Modules.BasicFinance.UI
{
    /// <summary>
    /// WPF 窗口与宿主主窗口关联辅助（规格 §20）。
    /// 通过 Win32 设置 Owner，保证窗口显示在 Excel/WPS 之上且不进入任务栏。
    /// </summary>
    internal static class WindowHelper
    {
        private const int GwlHwndParent = -8;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint flags);

        private static readonly IntPtr HwndTopmost = new IntPtr(-1);
        private const uint SwpNoSize = 0x0001;
        private const uint SwpNoMove = 0x0002;
        private const uint SwpNoActivate = 0x0010;

        public static void AttachOwner(Window window, IntPtr ownerHandle)
        {
            if (window == null || ownerHandle == IntPtr.Zero)
            {
                return;
            }

            window.SourceInitialized += (_, __) =>
            {
                var handle = new WindowInteropHelper(window).Handle;
                if (handle == IntPtr.Zero)
                {
                    return;
                }

                try
                {
                    if (Environment.Is64BitProcess)
                    {
                        SetWindowLongPtr(handle, GwlHwndParent, ownerHandle);
                    }
                    else
                    {
                        SetWindowLongPtr(handle, GwlHwndParent, ownerHandle);
                    }
                }
                catch
                {
                    // Owner 关联失败不影响窗口本身可用。
                }
            };
        }

        /// <summary>
        /// 将窗口置于宿主之上（不激活）。
        /// </summary>
        public static void BringToFront(Window window)
        {
            if (window == null)
            {
                return;
            }

            try
            {
                var handle = new WindowInteropHelper(window).Handle;
                if (handle != IntPtr.Zero)
                {
                    SetWindowPos(handle, HwndTopmost, 0, 0, 0, 0, SwpNoMove | SwpNoSize | SwpNoActivate);
                }
            }
            catch
            {
                // 忽略。
            }
        }
    }
}
