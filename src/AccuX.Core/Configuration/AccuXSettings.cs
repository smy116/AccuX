using System;

namespace AccuX.Core.Configuration
{
    /// <summary>
    /// AccuX 的应用级设置。对应 config.json 中的 "settings" 配置节。
    /// </summary>
    public sealed class AccuXSettings
    {
        public const long DefaultLargeSelectionWarning = 100000;
        public const long DefaultMaxProcessCells = 500000;
        public const bool DefaultAutoCheckForUpdates = true;

        /// <summary>超过该单元格数时提示用户确认。</summary>
        public long LargeSelectionWarning { get; set; } = DefaultLargeSelectionWarning;

        /// <summary>超过该单元格数时直接拒绝处理。</summary>
        public long MaxProcessCells { get; set; } = DefaultMaxProcessCells;

        /// <summary>是否在插件启动后自动检查稳定版升级。</summary>
        public bool AutoCheckForUpdates { get; set; } = DefaultAutoCheckForUpdates;

        /// <summary>上一次自动或手动检查的 UTC 时间。</summary>
        public DateTime? LastUpdateCheckUtc { get; set; }

        public static AccuXSettings CreateDefault()
        {
            return new AccuXSettings();
        }

        public AccuXSettings Clone()
        {
            return new AccuXSettings
            {
                LargeSelectionWarning = LargeSelectionWarning,
                MaxProcessCells = MaxProcessCells,
                AutoCheckForUpdates = AutoCheckForUpdates,
                LastUpdateCheckUtc = LastUpdateCheckUtc
            };
        }

        public bool TryValidate(out string message)
        {
            if (LargeSelectionWarning <= 0)
            {
                message = "警告单元格数量必须是大于 0 的整数。";
                return false;
            }

            if (MaxProcessCells <= 0)
            {
                message = "最大单元格数量必须是大于 0 的整数。";
                return false;
            }

            if (LargeSelectionWarning > MaxProcessCells)
            {
                message = "警告单元格数量不能大于最大单元格数量。";
                return false;
            }

            message = null;
            return true;
        }
    }
}
