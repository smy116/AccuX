using System;
using System.Security.Cryptography;
using System.Text;

namespace AccuX.Core.Configuration
{
    /// <summary>
    /// 基于 DPAPI（当前用户作用域）的敏感数据保护实现。
    /// 返回值为 Base64，便于写入只接受文本的宿主字段。
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
