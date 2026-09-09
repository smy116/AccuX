namespace AccuX.Core.Configuration
{
    /// <summary>
    /// 配置访问抽象。V1 统一由 <see cref="JsonConfigManager"/> 实现（见编码规则 25）。
    /// </summary>
    public interface IConfigManager
    {
        /// <summary>
        /// 读取指定配置节；不存在时返回 <paramref name="defaultValue"/>。
        /// </summary>
        T GetSection<T>(string sectionName, T defaultValue = default) where T : class, new();

        /// <summary>
        /// 当前配置文件路径。
        /// </summary>
        string ConfigFilePath { get; }

        /// <summary>
        /// 重新从磁盘加载配置。
        /// </summary>
        void Reload();
    }
}
