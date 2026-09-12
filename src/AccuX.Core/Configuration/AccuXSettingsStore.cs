using System;
using AccuX.Core.Logging;

namespace AccuX.Core.Configuration
{
    /// <summary>
    /// 应用级设置的加载、迁移和保存入口。
    /// </summary>
    public sealed class AccuXSettingsStore
    {
        private readonly object _gate = new object();
        private readonly IConfigManager _config;
        private readonly ILogger _logger;
        private AccuXSettings _current;

        public AccuXSettingsStore(IConfigManager config, ILogger logger = null)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
            _logger = logger ?? NullLogger.Instance;
        }

        public string ConfigFilePath
        {
            get { return _config.ConfigFilePath; }
        }

        public AccuXSettings Load()
        {
            lock (_gate)
            {
                if (_current != null)
                {
                    return _current.Clone();
                }

                var migrated = false;
                var needsPersist = false;
                var missingSettings = new AccuXSettings
                {
                    LargeSelectionWarning = long.MinValue,
                    MaxProcessCells = long.MinValue,
                    AutoCheckForUpdates = false
                };
                var settings = _config.GetSection("settings", missingSettings);
                if (IsMissingMarker(settings))
                {
                    settings = MigrateLegacySettings();
                    migrated = true;
                }
                else
                {
                    if (settings == null)
                    {
                        settings = AccuXSettings.CreateDefault();
                    }
                    if (!Normalize(settings))
                    {
                        _logger.Warn("应用设置无效，已回退到默认值。");
                        needsPersist = true;
                    }
                }

                _current = settings;
                if (migrated || needsPersist)
                {
                    TryPersist(settings, migrated ? "已迁移旧版单元格阈值到 settings 配置节。" : "已保存规范化的应用设置。");
                }

                return _current.Clone();
            }
        }

        public AccuXSettings Current
        {
            get
            {
                return Load();
            }
        }

        /// <summary>
        /// 只保存设置窗口可编辑的字段，避免覆盖后台检查刚写入的时间戳。
        /// </summary>
        public AccuXSettings SaveEditable(AccuXSettings value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            if (!value.TryValidate(out var message))
            {
                throw new ArgumentException(message, nameof(value));
            }

            lock (_gate)
            {
                var current = Load();
                current.LargeSelectionWarning = value.LargeSelectionWarning;
                current.MaxProcessCells = value.MaxProcessCells;
                current.AutoCheckForUpdates = value.AutoCheckForUpdates;
                Persist(current);
                _current = current;
                return current.Clone();
            }
        }

        public AccuXSettings RecordUpdateCheck(DateTime utcNow)
        {
            lock (_gate)
            {
                var current = Load();
                current.LastUpdateCheckUtc = utcNow.ToUniversalTime();
                Persist(current);
                _current = current;
                return current.Clone();
            }
        }

        private AccuXSettings MigrateLegacySettings()
        {
            var basic = _config.GetSection("basicFinance", new LegacyThresholds());
            if (IsValidPair(basic))
            {
                return CreateFromLegacy(basic);
            }

            var compare = _config.GetSection("compare", new LegacyThresholds());
            if (IsValidPair(compare))
            {
                return CreateFromLegacy(compare);
            }

            return AccuXSettings.CreateDefault();
        }

        private static AccuXSettings CreateFromLegacy(LegacyThresholds legacy)
        {
            return new AccuXSettings
            {
                LargeSelectionWarning = legacy.LargeSelectionWarning,
                MaxProcessCells = legacy.MaxProcessCells,
                AutoCheckForUpdates = AccuXSettings.DefaultAutoCheckForUpdates
            };
        }

        private static bool IsValidPair(LegacyThresholds value)
        {
            return value != null
                && value.LargeSelectionWarning > 0
                && value.MaxProcessCells > 0
                && value.LargeSelectionWarning <= value.MaxProcessCells;
        }

        private static bool Normalize(AccuXSettings settings)
        {
            if (settings == null)
            {
                return false;
            }

            if (settings.TryValidate(out _))
            {
                return true;
            }

            settings.LargeSelectionWarning = AccuXSettings.DefaultLargeSelectionWarning;
            settings.MaxProcessCells = AccuXSettings.DefaultMaxProcessCells;
            return false;
        }

        private static bool IsMissingMarker(AccuXSettings settings)
        {
            return settings != null
                && settings.LargeSelectionWarning == long.MinValue
                && settings.MaxProcessCells == long.MinValue
                && !settings.AutoCheckForUpdates;
        }

        private void TryPersist(AccuXSettings settings, string message)
        {
            try
            {
                Persist(settings);
                _logger.Info(message);
            }
            catch (Exception ex)
            {
                _logger.Warn("应用设置写入失败：" + ex.Message);
            }
        }

        private void Persist(AccuXSettings settings)
        {
            _config.SaveSection("settings", settings);
        }

        private sealed class LegacyThresholds
        {
            public long LargeSelectionWarning { get; set; }

            public long MaxProcessCells { get; set; }
        }
    }
}
