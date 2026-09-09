using System;
using System.Globalization;

namespace AccuX.Modules.BasicFinance.Comment
{
    /// <summary>
    /// 批注文本长度校验。按 Unicode 文本元素计数，CRLF 作为一个换行字符。
    /// </summary>
    public static class CommentTextValidator
    {
        public const int MaxLength = 1000;

        public static int CountTextElements(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            var normalized = NormalizeLineEndings(text);
            return new StringInfo(normalized).LengthInTextElements;
        }

        public static bool IsWithinLimit(string text)
        {
            return CountTextElements(text) <= MaxLength;
        }

        public static string NormalizeLineEndings(string text)
        {
            return (text ?? string.Empty).Replace("\r\n", "\n").Replace("\r", "\n");
        }

        public static string BuildCounter(string text)
        {
            return string.Format(
                CultureInfo.CurrentCulture,
                "当前{0}/{1}字",
                CountTextElements(text),
                MaxLength);
        }
    }
}
