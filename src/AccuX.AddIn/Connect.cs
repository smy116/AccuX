using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AccuX.AddIn.ComInterop;
using AccuX.AddIn.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace AccuX.AddIn
{
    /// <summary>
    /// AccuX COM Add-in 入口与 Composition Root（规格 §5.1 / §30）。
    /// <para>
    /// 负责：COM 生命周期、Ribbon 加载与回调、初始化 Host/Config/Logger、
    /// 创建 Pipeline 并注入 Host 实现、显式注册模块、分发 Command。
    /// 不包含任何财务算法。
    /// </para>
    /// </summary>
    [ComVisible(true)]
    [Guid("7C1F0E4A-9B2D-4E8C-A1F3-6D5E8B0C2A11")]
    [ProgId("AccuX.AddIn.Connect")]
    // 必须是 AutoDual：Office 通过 IDispatch 按名称调用 Ribbon 回调
    // （onLoad / onAction / getImage）。ClassInterfaceType.None 不会生成类接口，
    // 这些回调无法被 IDispatch 解析，Office 会静默放弃加载 Ribbon（选项卡不显示）。
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class Connect : IDTExtensibility2, IRibbonExtensibility, IDisposable
    {
        private Excel.Application _application;
        private AddInCompositionRoot _compositionRoot;
        private bool _disposed;

        /// <summary>
        /// COM 连接建立。
        /// </summary>
        public void OnConnection(
            object application,
            ext_ConnectMode connectMode,
            object addInInst,
            ref Array custom)
        {
            try
            {
                _application = application as Excel.Application;
                if (_application == null)
                {
                    return;
                }

                _compositionRoot = new AddInCompositionRoot(_application);
                _compositionRoot.Start();
            }
            catch (Exception ex)
            {
                // 入口异常绝不穿出 COM 边界。
                TryLogFatal("OnConnection 失败", ex);
            }
        }

        /// <summary>
        /// COM 断开连接。
        /// </summary>
        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                _compositionRoot?.Stop();
                _compositionRoot = null;
            }
            catch (Exception ex)
            {
                TryLogFatal("OnDisconnection 失败", ex);
            }
            finally
            {
                _application = null;
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
            try
            {
                _compositionRoot?.Stop();
            }
            catch
            {
                // 关闭阶段忽略异常。
            }
        }

        /// <summary>
        /// 返回嵌入的 Office CustomUI XML。
        /// </summary>
        public string GetCustomUI(string ribbonId)
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var name = "AccuX.AddIn.Ribbon.AccuXRibbon.xml";
                using (var stream = assembly.GetManifestResourceStream(name))
                {
                    if (stream == null)
                    {
                        _compositionRoot?.Logger?.Error("未找到嵌入资源: " + name);
                        return string.Empty;
                    }

                    using (var reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                TryLogFatal("GetCustomUI 失败", ex);
                return string.Empty;
            }
        }

        /// <summary>
        /// Ribbon 加载完成回调。
        /// </summary>
        public void OnRibbonLoad(object ribbon)
        {
        }

        /// <summary>
        /// 所有 Ribbon 按钮统一进入 CommandDispatcher（规格 §7）。
        /// </summary>
        public void OnAction(IRibbonControl control)
        {
            try
            {
                if (_compositionRoot == null)
                {
                    return;
                }

                var result = _compositionRoot.Dispatch(control?.Id);
                if (result != null && result.ShowMessage && !string.IsNullOrEmpty(result.Message))
                {
                    if (result.Success)
                    {
                        _compositionRoot.Prompt.ShowMessage(result.Message);
                    }
                    else
                    {
                        _compositionRoot.Prompt.ShowError(result.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                // 任何异常都不得穿出 COM callback。
                TryLogFatal("OnAction 失败", ex);
            }
        }

        /// <summary>
        /// Ribbon 图标回调。
        /// </summary>
        public object GetImage(IRibbonControl control)
        {
            return RibbonImageProvider.GetImage(control?.Id);
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _compositionRoot?.Stop();
        }

        private void TryLogFatal(string message, Exception exception)
        {
            try
            {
                _compositionRoot?.Logger.Error(message, exception);
            }
            catch
            {
                // 最后兜底：日志本身失败也不能抛出。
            }
        }
    }
}
