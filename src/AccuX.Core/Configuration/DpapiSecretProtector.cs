using System;
using System.Security.Cryptography;
using System.Text;

namespace AccuX.Core.Configuration
{
    /// <summary>
    /// 基于 DPAPI（当前用户作用域）的敏感配置保护实现。
    /// V1 无实际使用场景，保留以便后续扩展。
    /// </summary>
    public sealed class DpapiSecretProtector : ISecretProtector
    {
        public string Protect(string plainText)
        {
            if (string.IsNullOrEmpty(plainText))
            {
                return plainText;
            }

            var bytes = Encoding.UTF8.GetBytes(plainText);
            var protectedBytes = ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(protectedBytes);
        }

        public string Unprotect(string protectedText)
        {
            if (string.IsNullOrEmpty(protectedText))
            {
                return protectedText;
            }

            var protectedBytes = Convert.FromBase64String(protectedText);
            var bytes = ProtectedData.Unprotect(protectedBytes, null, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(bytes);
        }
    }
}
