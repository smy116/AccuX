using System;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;

namespace AccuX.Core.Configuration
{
    /// <summary>
    /// 基于 JSON 的配置管理器。
    /// 默认路径为 %AppData%\AccuX\config.json；文件不存在时使用传入的默认值，不抛异常。
    /// </summary>
    public sealed class JsonConfigManager : IConfigManager
    {
        private readonly object _gate = new object();
        private readonly JsonSerializerSettings _serializerSettings;
        private JObject _root;

        public JsonConfigManager(string configFilePath)
        {
            if (string.IsNullOrWhiteSpace(configFilePath))
            {
                throw new ArgumentException("配置文件路径不能为空。", nameof(configFilePath));
            }

            ConfigFilePath = configFilePath;
            _serializerSettings = new JsonSerializerSettings
            {
                MissingMemberHandling = MissingMemberHandling.Ignore,
                ObjectCreationHandling = ObjectCreationHandling.Replace,
                NullValueHandling = NullValueHandling.Ignore,
                ContractResolver = new DefaultContractResolver
                {
                    NamingStrategy = new CamelCaseNamingStrategy()
                }
            };

            Reload();
        }

        public static string DefaultConfigPath
        {
            get
            {
                var root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                return Path.Combine(root, "AccuX", "config.json");
            }
        }

        public string ConfigFilePath { get; }

        public void Reload()
        {
            lock (_gate)
            {
                _root = LoadRoot();
            }
        }

        public T GetSection<T>(string sectionName, T defaultValue = default) where T : class, new()
        {
            if (string.IsNullOrWhiteSpace(sectionName))
            {
                throw new ArgumentException("配置节名称不能为空。", nameof(sectionName));
            }

            try
            {
                JToken section;
                lock (_gate)
                {
                    section = _root?[sectionName];
                }

                if (section == null || section.Type == JTokenType.Null)
                {
                    return defaultValue ?? new T();
                }

                var value = section.ToObject<T>(JsonSerializer.Create(_serializerSettings));
                return value ?? defaultValue ?? new T();
            }
            catch
            {
                // 配置损坏时回退到默认值，保证插件仍可启动。
                return defaultValue ?? new T();
            }
        }

        /// <summary>
        /// 将配置节写回磁盘；用于设置窗口保存。
        /// </summary>
        public void SaveSection<T>(string sectionName, T value) where T : class, new()
        {
            if (string.IsNullOrWhiteSpace(sectionName))
            {
                throw new ArgumentException("配置节名称不能为空。", nameof(sectionName));
            }

            lock (_gate)
            {
                _root = _root ?? new JObject();
                var serializer = JsonSerializer.Create(_serializerSettings);
                // 保存配置节时保留 null 字段，确保设置结构中的 lastUpdateCheckUtc
                // 在尚未检测时也明确写出；读取仍沿用原有的忽略 null 行为。
                serializer.NullValueHandling = NullValueHandling.Include;
                _root[sectionName] = value == null
                    ? JValue.CreateNull()
                    : JToken.FromObject(value, serializer);

                var directory = Path.GetDirectoryName(ConfigFilePath);
                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                File.WriteAllText(ConfigFilePath, _root.ToString(Formatting.Indented), new UTF8Encoding(false));
            }
        }

        private JObject LoadRoot()
        {
            try
            {
                if (!File.Exists(ConfigFilePath))
                {
                    return new JObject();
                }

                var text = File.ReadAllText(ConfigFilePath, Encoding.UTF8);
                if (string.IsNullOrWhiteSpace(text))
                {
                    return new JObject();
                }

                return JObject.Parse(text);
            }
            catch
            {
                return new JObject();
            }
        }
    }
}
