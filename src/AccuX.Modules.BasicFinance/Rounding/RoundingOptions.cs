using System;
using System.Globalization;

namespace AccuX.Modules.BasicFinance.Rounding
{
    /// <summary>
    /// 一键舍入选项（规格 §12）。
    /// </summary>
    public sealed class RoundingOptions
    {
        /// <summary>保留小数位数；默认 2 位。</summary>
        public int Digits { get; set; } = 2;

        /// <summary>允许的最大小数位，防止异常输入。</summary>
        public const int MaxDigits = 15;

        public RoundingOptions()
        {
        }

        public RoundingOptions(int digits)
        {
            Digits = digits;
        }

        public bool IsValid
        {
            get { return Digits >= 0 && Digits <= MaxDigits; }
        }

        /// <summary>
        /// 校验并返回规范化后的选项。
        /// </summary>
        public static bool TryNormalize(int digits, out RoundingOptions options, out string error)
        {
            options = null;
            error = null;

            if (digits < 0 || digits > MaxDigits)
            {
                error = string.Format(CultureInfo.CurrentCulture, "小数位数必须在 0 到 {0} 之间。", MaxDigits);
                return false;
            }

            options = new RoundingOptions(digits);
            return true;
        }
    }
}
