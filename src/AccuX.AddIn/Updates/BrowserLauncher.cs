using System;
using System.Diagnostics;

namespace AccuX.AddIn.Updates
{
    internal interface IBrowserLauncher
    {
        bool TryOpen(string url, out string errorMessage);
    }

    internal sealed class BrowserLauncher : IBrowserLauncher
    {
        public bool TryOpen(string url, out string errorMessage)
        {
            errorMessage = null;
            if (string.IsNullOrWhiteSpace(url)
                || !Uri.TryCreate(url, UriKind.Absolute, out var uri)
                || !string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
            {
                errorMessage = "升级下载地址无效。";
                return false;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = uri.AbsoluteUri,
                    UseShellExecute = true
                });
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = "无法打开升级下载地址：" + ex.Message;
                return false;
            }
        }
    }
}
