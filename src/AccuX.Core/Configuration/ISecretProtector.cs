namespace AccuX.Core.Configuration
{
    /// <summary>
    /// 敏感数据保护抽象（默认实现为当前用户 DPAPI）。
    /// </summary>
    public interface ISecretProtector
    {
        string Protect(string plainText);

        string Unprotect(string protectedText);
    }
}
