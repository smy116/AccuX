using System;
using System.Diagnostics;
using System.IO;

namespace AccuX.AddIn.Updates
{
    internal interface IInstallerLauncher
    {
        bool TryLaunch(string installerPath, out string errorMessage);
    }

    internal sealed class InstallerLauncher : IInstallerLauncher
    {
        public bool TryLaunch(string installerPath, out string errorMessage)
        {
            errorMessage = null;
            if (string.IsNullOrWhiteSpace(installerPath) || !File.Exists(installerPath))
            {
                errorMessage = "找不到已下载的升级安装包。";
                return false;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = installerPath,
                    WorkingDirectory = Path.GetDirectoryName(installerPath),
                    UseShellExecute = true
                });
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = "无法启动升级安装包：" + ex.Message;
                return false;
            }
        }
    }
}
