using System;
using System.IO;
using AccuX.Core.Configuration;
using Xunit;

namespace AccuX.Core.Tests
{
    public class JsonConfigManagerTests : IDisposable
    {
        private readonly string _path;

        public JsonConfigManagerTests()
        {
            _path = Path.Combine(Path.GetTempPath(), "accux-test-" + Guid.NewGuid() + ".json");
        }

        public void Dispose()
        {
            try
            {
                if (File.Exists(_path))
                {
                    File.Delete(_path);
                }
            }
            catch
            {
                // 忽略清理失败。
            }
        }

        [Fact]
        public void GetSection_FileMissing_ReturnsDefault()
        {
            var manager = new JsonConfigManager(_path);

            var section = manager.GetSection("basicFinance", new TestSection { Value = 7 });

            Assert.Equal(7, section.Value);
        }

        [Fact]
        public void GetSection_ReadsJson()
        {
            File.WriteAllText(_path, "{ \"basicFinance\": { \"value\": 42 } }");

            var manager = new JsonConfigManager(_path);
            var section = manager.GetSection("basicFinance", new TestSection());

            Assert.Equal(42, section.Value);
        }

        [Fact]
        public void GetSection_CorruptJson_FallsBackToDefault()
        {
            File.WriteAllText(_path, "{ this is not json");

            var manager = new JsonConfigManager(_path);
            var section = manager.GetSection("basicFinance", new TestSection { Value = 3 });

            Assert.Equal(3, section.Value);
        }

        [Fact]
        public void SaveSection_RoundTrips()
        {
            var manager = new JsonConfigManager(_path);
            manager.SaveSection("basicFinance", new TestSection { Value = 99 });

            var reloaded = new JsonConfigManager(_path);
            var section = reloaded.GetSection("basicFinance", new TestSection());

            Assert.Equal(99, section.Value);
        }

        [Fact]
        public void SaveSection_PreservesOtherSectionsAndUsesCamelCase()
        {
            File.WriteAllText(_path, "{\"otherSection\":{\"keepMe\":true}}");

            var manager = new JsonConfigManager(_path);
            manager.SaveSection("settings", new NamedSection { SomeValue = 12 });

            var json = File.ReadAllText(_path);
            Assert.Contains("\"otherSection\"", json);
            Assert.Contains("\"keepMe\": true", json);
            Assert.Contains("\"someValue\": 12", json);
            Assert.DoesNotContain("SomeValue", json);
        }

        [Fact]
        public void Reload_PicksUpExternalChanges()
        {
            var manager = new JsonConfigManager(_path);
            File.WriteAllText(_path, "{ \"basicFinance\": { \"value\": 5 } }");

            manager.Reload();
            var section = manager.GetSection("basicFinance", new TestSection());

            Assert.Equal(5, section.Value);
        }

        private sealed class TestSection
        {
            public int Value { get; set; }
        }

        private sealed class NamedSection
        {
            public int SomeValue { get; set; }
        }
    }
}
