using System;
using System.Linq;
using System.Text;
using AccuX.Core.Configuration;
using AccuX.Core.Operations;

namespace AccuX.Modules.BasicFinance.Comment
{
    /// <summary>
    /// 普通/加密批注编解码器。
    /// 加密载荷格式为 [固定标记][版本号][DPAPI 密文字节]，整体再 Base64 编码。
    /// </summary>
    public sealed class CommentCodec
    {
        private static readonly byte[] Magic = Encoding.ASCII.GetBytes("AccuX.Comment");
        private const byte CurrentVersion = 1;

        private readonly ISecretProtector _protector;

        public CommentCodec(ISecretProtector protector = null)
        {
            _protector = protector ?? new DpapiSecretProtector();
        }

        public string Encode(string text, CommentKind kind)
        {
            text = text ?? string.Empty;
            if (kind == CommentKind.Plain)
            {
                return text;
            }

            if (kind != CommentKind.Encrypted)
            {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            var protectedBase64 = _protector.Protect(text);
            if (string.IsNullOrEmpty(protectedBase64))
            {
                throw new InvalidOperationException("加密批注失败：保护器未返回有效内容。");
            }

            byte[] protectedBytes;
            try
            {
                protectedBytes = Convert.FromBase64String(protectedBase64);
            }
            catch (FormatException ex)
            {
                throw new InvalidOperationException("加密批注失败：保护器返回的内容不是有效 Base64。", ex);
            }

            var payload = new byte[Magic.Length + 1 + protectedBytes.Length];
            Buffer.BlockCopy(Magic, 0, payload, 0, Magic.Length);
            payload[Magic.Length] = CurrentVersion;
            Buffer.BlockCopy(protectedBytes, 0, payload, Magic.Length + 1, protectedBytes.Length);
            return Convert.ToBase64String(payload);
        }

        public CommentContent Decode(CellComment comment)
        {
            if (comment == null || !comment.Exists)
            {
                return CommentContent.Empty;
            }

            return Decode(comment.Text);
        }

        public CommentContent Decode(string storedText)
        {
            storedText = storedText ?? string.Empty;
            byte[] payload;
            try
            {
                payload = Convert.FromBase64String(storedText);
            }
            catch (FormatException)
            {
                return CommentContent.Plain(storedText);
            }

            if (!HasMagic(payload))
            {
                // 普通批注即使恰好是 Base64 文本，也不能仅凭 Base64 形状判定为加密批注。
                return CommentContent.Plain(storedText);
            }

            // 只有完整的“标记 + 版本号”才进入加密格式解析；普通文本恰好解码为标记本身时仍按普通批注处理。
            if (payload.Length == Magic.Length)
            {
                return CommentContent.Plain(storedText);
            }

            if (payload.Length < Magic.Length)
            {
                return CommentContent.Plain(storedText);
            }

            var version = payload[Magic.Length];
            if (version != CurrentVersion)
            {
                return CommentContent.Unreadable("加密批注版本不受支持，无法解密。", true);
            }

            var protectedLength = payload.Length - Magic.Length - 1;
            if (protectedLength <= 0)
            {
                return CommentContent.Unreadable("加密批注内容为空或已损坏，无法解密。", true);
            }

            var protectedBytes = new byte[protectedLength];
            Buffer.BlockCopy(payload, Magic.Length + 1, protectedBytes, 0, protectedLength);
            try
            {
                var plainText = _protector.Unprotect(Convert.ToBase64String(protectedBytes));
                return CommentContent.Encrypted(plainText ?? string.Empty);
            }
            catch (Exception ex)
            {
                return CommentContent.Unreadable(
                    "加密批注无法解密，当前 Windows 用户可能与写入批注的用户不同，或批注已损坏。",
                    true,
                    ex);
            }
        }

        private static bool HasMagic(byte[] payload)
        {
            if (payload == null || payload.Length < Magic.Length)
            {
                return false;
            }

            return Magic.SequenceEqual(payload.Take(Magic.Length));
        }
    }

    /// <summary>编解码后的批注编辑状态。</summary>
    public sealed class CommentContent
    {
        private CommentContent(
            bool exists,
            bool encrypted,
            bool unreadable,
            string text,
            string errorMessage,
            Exception exception)
        {
            Exists = exists;
            IsEncrypted = encrypted;
            IsUnreadable = unreadable;
            Text = text ?? string.Empty;
            ErrorMessage = errorMessage ?? string.Empty;
            Exception = exception;
        }

        public bool Exists { get; }

        public bool IsEncrypted { get; }

        public bool IsUnreadable { get; }

        public string Text { get; }

        public string ErrorMessage { get; }

        public Exception Exception { get; }

        public static CommentContent Empty
        {
            get { return new CommentContent(false, false, false, string.Empty, string.Empty, null); }
        }

        public static CommentContent Plain(string text)
        {
            return new CommentContent(true, false, false, text, string.Empty, null);
        }

        public static CommentContent Encrypted(string text)
        {
            return new CommentContent(true, true, false, text, string.Empty, null);
        }

        public static CommentContent Unreadable(string message, bool encrypted, Exception exception = null)
        {
            return new CommentContent(true, encrypted, true, string.Empty, message, exception);
        }
    }
}
