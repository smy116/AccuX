using System;

namespace AccuX.Core.Operations
{
    /// <summary>
    /// 需要由 <see cref="IHostStateScope"/> 保存并恢复的宿主 Application 级状态（规格 §5.2）。
    /// 只影响本次操作真正需要修改的状态；未修改的状态不得主动改动。
    /// </summary>
    public sealed class HostStateOptions
    {
        /// <summary>需要临时关闭屏幕更新。</summary>
        public bool DisableScreenUpdating { get; set; } = true;

        /// <summary>需要临时切换计算模式为手动。</summary>
        public bool ManualCalculation { get; set; }

        /// <summary>需要临时关闭事件。</summary>
        public bool DisableEvents { get; set; } = true;

        /// <summary>需要临时关闭提示对话框。</summary>
        public bool DisableDisplayAlerts { get; set; } = true;

        /// <summary>可选的状态栏文本；null 表示不修改。</summary>
        public string StatusBarText { get; set; }

        public static HostStateOptions Default
        {
            get { return new HostStateOptions(); }
        }
    }

    /// <summary>
    /// 宿主 Application 状态作用域。
    /// <para>
    /// 职责仅限：保存原始状态 → 按需修改 → 操作结束后在 Dispose 中恢复。
    /// <b>不得</b>承担 Values/Formulas 备份、工作表数据恢复、AccuX Undo 或事务 Commit/Rollback。
    /// </para>
    /// </summary>
    public interface IHostStateScope : IDisposable
    {
    }
}
