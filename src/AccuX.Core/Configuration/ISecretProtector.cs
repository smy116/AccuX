namespace AccuX.Core.Configuration
{
    /// <summary>
    /// 敏感配置保护抽象（DPAPI）。
    /// V1 没有 API Key 等敏感配置，因此没有必须使用的场景，仅保留接口以备扩展。
    /// </summary>
    public interface ISecretProtector
    {
        string Protect(string plainText);

        string Unprotect(string protectedText);
    }
}
