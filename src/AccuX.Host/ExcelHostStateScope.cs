using System;
using AccuX.Core.Operations;
using Excel = Microsoft.Office.Interop.Excel;

namespace AccuX.Host
{
    /// <summary>
    /// 宿主 Application 状态作用域实现（规格 §5.2）。
    /// <para>
    /// 只保存并恢复本次真正修改过的状态；Dispose 中即使恢复失败也不得抛出异常。
    /// 不承担工作表数据恢复、AccuX Undo 或事务语义。
    /// </para>
    /// </summary>
    internal sealed class ExcelHostStateScope : IHostStateScope
    {
        private readonly Excel.Application _application;
        private readonly HostStateOptions _options;

        private bool _screenUpdatingSaved;
        private bool _screenUpdatingOriginal;

        private bool _calculationSaved;
        private Excel.XlCalculation _calculationOriginal;

        private bool _eventsSaved;
        private bool _eventsOriginal;

        private bool _alertsSaved;
        private bool _alertsOriginal;

        private bool _statusBarSaved;
        private object _statusBarOriginal;

        private bool _disposed;

        public ExcelHostStateScope(Excel.Application application, HostStateOptions options)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _options = options ?? HostStateOptions.Default;

            Apply();
        }

        private void Apply()
        {
            if (_options.DisableScreenUpdating)
            {
                try
                {
                    _screenUpdatingOriginal = _application.ScreenUpdating;
                    _screenUpdatingSaved = true;
                    _application.ScreenUpdating = false;
                }
                catch
                {
                    _screenUpdatingSaved = false;
                }
            }

            if (_options.ManualCalculation)
            {
                try
                {
                    _calculationOriginal = _application.Calculation;
                    _calculationSaved = true;
                    _application.Calculation = Excel.XlCalculation.xlCalculationManual;
                }
                catch
                {
                    _calculationSaved = false;
                }
            }

            if (_options.DisableEvents)
            {
                try
                {
                    _eventsOriginal = _application.EnableEvents;
                    _eventsSaved = true;
                    _application.EnableEvents = false;
                }
                catch
                {
                    _eventsSaved = false;
                }
            }

            if (_options.DisableDisplayAlerts)
            {
                try
                {
                    _alertsOriginal = _application.DisplayAlerts;
                    _alertsSaved = true;
                    _application.DisplayAlerts = false;
                }
                catch
                {
                    _alertsSaved = false;
                }
            }

            if (!string.IsNullOrEmpty(_options.StatusBarText))
            {
                try
                {
                    _statusBarOriginal = _application.StatusBar;
                    _statusBarSaved = true;
                    _application.StatusBar = _options.StatusBarText;
                }
                catch
                {
                    _statusBarSaved = false;
                }
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;

            // 逆序恢复，任何一步失败都不得影响后续恢复。
            if (_statusBarSaved)
            {
                Try(() => _application.StatusBar = _statusBarOriginal);
            }

            if (_alertsSaved)
            {
                Try(() => _application.DisplayAlerts = _alertsOriginal);
            }

            if (_eventsSaved)
            {
                Try(() => _application.EnableEvents = _eventsOriginal);
            }

            if (_calculationSaved)
            {
                Try(() => _application.Calculation = _calculationOriginal);
            }

            if (_screenUpdatingSaved)
            {
                Try(() => _application.ScreenUpdating = _screenUpdatingOriginal);
            }
        }

        private static void Try(Action action)
        {
            try
            {
                action();
            }
            catch
            {
                // 恢复失败不得抛出，避免掩盖原始业务异常。
            }
        }
    }
}
