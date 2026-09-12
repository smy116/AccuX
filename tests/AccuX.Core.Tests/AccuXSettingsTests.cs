using System;
using System.IO;
using AccuX.Core.Configuration;
using Xunit;

namespace AccuX.Core.Tests
{
    public sealed class AccuXSettingsTests : IDisposable
    {
        private readonly string _path = Path.Combine(Path.GetTempPath(), "accux-settings-" + Guid.NewGuid() + ".json");

        public void Dispose()
        {
            try
            {
                if (File.Exists(_path)) File.Delete(_path);
            }
            catch
            {
                // 测试清理失败不影响断言结果。
            }
        }

        [Fact]
        public void Store_MigratesBasicFinanceBeforeCompare()
        {
            File.WriteAllText(_path,
                "{\"basicFinance\":{\"largeSelectionWarning\":12,\"maxProcessCells\":34},"
                + "\"compare\":{\"largeSelectionWarning\":56,\"maxProcessCells\":78}}" );

            var manager = new JsonConfigManager(_path);
            var settings = new AccuXSettingsStore(manager).Load();

            Assert.Equal(12, settings.LargeSelectionWarning);
            Assert.Equal(34, settings.MaxProcessCells);
            var saved = File.ReadAllText(_path);
            Assert.Contains("\"settings\"", saved);
            Assert.Contains("largeSelectionWarning", saved);
            Assert.Contains("\"lastUpdateCheckUtc\": null", saved);
        }

        [Fact]
        public void Store_FallsBackToCompareWhenBasicFinanceIsInvalid()
        {
            File.WriteAllText(_path,
                "{\"basicFinance\":{\"largeSelectionWarning\":0,\"maxProcessCells\":34},"
                + "\"compare\":{\"largeSelectionWarning\":56,\"maxProcessCells\":78}}" );

            var settings = new AccuXSettingsStore(new JsonConfigManager(_path)).Load();

            Assert.Equal(56, settings.LargeSelectionWarning);
            Assert.Equal(78, settings.MaxProcessCells);
        }

        [Fact]
        public void Store_InvalidCanonicalSettingsUseDefaults()
        {
            File.WriteAllText(_path,
                "{\"settings\":{\"largeSelectionWarning\":900,\"maxProcessCells\":100}}" );

            var manager = new JsonConfigManager(_path);
            var settings = new AccuXSettingsStore(manager).Load();

            Assert.Equal(AccuXSettings.DefaultLargeSelectionWarning, settings.LargeSelectionWarning);
            Assert.Equal(AccuXSettings.DefaultMaxProcessCells, settings.MaxProcessCells);
            Assert.Equal(AccuXSettings.DefaultAutoCheckForUpdates, settings.AutoCheckForUpdates);
        }

        [Fact]
        public void Store_SaveEditablePreservesUpdateCheckTime()
        {
            var store = new AccuXSettingsStore(new JsonConfigManager(_path));
            var checkedAt = new DateTime(2026, 9, 12, 1, 2, 3, DateTimeKind.Utc);
            store.RecordUpdateCheck(checkedAt);

            var updated = store.SaveEditable(new AccuXSettings
            {
                LargeSelectionWarning = 20,
                MaxProcessCells = 30,
                AutoCheckForUpdates = false
            });

            Assert.Equal(20, updated.LargeSelectionWarning);
            Assert.Equal(30, updated.MaxProcessCells);
            Assert.False(updated.AutoCheckForUpdates);
            Assert.Equal(checkedAt, updated.LastUpdateCheckUtc);
        }

        [Fact]
        public void Settings_RejectWarningAboveMaximum()
        {
            var settings = new AccuXSettings
            {
                LargeSelectionWarning = 31,
                MaxProcessCells = 30
            };

            Assert.False(settings.TryValidate(out var message));
            Assert.Contains("不能大于", message);
        }
    }
}
