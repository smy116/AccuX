using System;
using System.Collections.Generic;
using AccuX.Core.Commands;
using AccuX.Core.Modules;
using AccuX.Core.Operations;

namespace AccuX.Modules.Compare
{
    public sealed class RegionCompareModule : IAccuXModule
    {
        public const string ModuleId = "AccuX.Modules.Compare";

        private IAccuXContext _context;
        private RegionCompareWindow _window;
        private CommandDefinition[] _commands = Array.Empty<CommandDefinition>();

        public string Id { get { return ModuleId; } }

        public void Initialize(IAccuXContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _commands = new[]
            {
                new CommandDefinition(
                    "accux.compare.exists",
                    "存在对比",
                    ModuleId,
                    Execute,
                    "比较两个区域中出现过的单元格值，并按区域独有与相同项分类。",
                    "accux.compare.exists")
            };
            context.Logger.Info("Compare 模块初始化完成。");
        }

        public IEnumerable<CommandDefinition> GetCommands()
        {
            return _commands;
        }

        public void Shutdown()
        {
            try { _window?.Close(); } catch { }
            try { _context?.RegionCompareHost?.ReleaseSession(); } catch { }
            _window = null;
            _commands = Array.Empty<CommandDefinition>();
        }

        private CommandResult Execute(CommandExecutionContext execution)
        {
            var host = _context.RegionCompareHost;
            if (host == null)
            {
                return CommandResult.Failed("当前宿主不支持区域对比功能。");
            }

            if (_window != null)
            {
                _window.Activate();
                return CommandResult.Cancelled();
            }

            var config = _context.Config.GetSection("compare", new CompareConfig());
            _window = new RegionCompareWindow(
                host,
                new RegionCompareService(),
                config,
                _context.Logger,
                _context.Host?.MainWindowHandle);
            _window.Closed += (sender, args) =>
            {
                try { _context.RegionCompareHost?.ReleaseSession(); } catch { }
                _window = null;
            };
            try
            {
                _window.Show();
                return CommandResult.Cancelled();
            }
            catch (Exception ex)
            {
                _window = null;
                return CommandResult.Failed("无法打开区域对比窗口：" + ex.Message, ex);
            }
        }
    }
}
