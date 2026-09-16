using System;
using System.Collections.Generic;
using AccuX.Core.Commands;
using AccuX.Core.Logging;

namespace AccuX.Core.Modules
{
    /// <summary>
    /// 模块注册表（规格 §6）。
    /// <para>
    /// AccuX 只保存 AddIn 显式传入的模块实例与命令注册结果：
    /// <b>不</b>扫描 DLL、<b>不</b>反射发现、<b>不</b>热加载、<b>不</b>解析依赖。
    /// 单个模块初始化失败会被捕获并记录，不影响其他模块与整个 Add-in。
    /// </para>
    /// </summary>
    public sealed class ModuleRegistry : IDisposable
    {
        private readonly List<IAccuXModule> _modules = new List<IAccuXModule>();
        private readonly ILogger _logger;
        private bool _disposed;

        public ModuleRegistry(ILogger logger)
        {
            _logger = logger ?? NullLogger.Instance;
        }

        public IReadOnlyList<IAccuXModule> Modules
        {
            get { return _modules; }
        }

        /// <summary>
        /// 注册并初始化模块，同时把其命令注册到 dispatcher。
        /// </summary>
        public void Register(IAccuXModule module, IAccuXContext context, CommandDispatcher dispatcher)
        {
            if (module == null)
            {
                throw new ArgumentNullException(nameof(module));
            }

            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            if (dispatcher == null)
            {
                throw new ArgumentNullException(nameof(dispatcher));
            }

            try
            {
                module.Initialize(context);
                dispatcher.RegisterRange(module.GetCommands());
                _modules.Add(module);
                _logger.Info("模块已加载: " + module.Id);
            }
            catch (Exception ex)
            {
                // 单个模块失败不得导致 AccuX 整体失效。
                _logger.Error("模块初始化失败，已隔离: " + module.Id, ex);
            }
        }

        public void Shutdown()
        {
            foreach (var module in _modules)
            {
                try
                {
                    module.Shutdown();
                }
                catch (Exception ex)
                {
                    _logger.Error("模块卸载失败: " + module.Id, ex);
                }
            }

            _modules.Clear();
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            Shutdown();
        }
    }
}
