using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace AccuX.Modules.BasicFinance.ChineseAmount
{
    /// <summary>
    /// 金额转中文大写（规格 §14）。
    /// <para>纯 C# 算法，可脱离 Excel/WPS 单元测试。</para>
    /// <para>
    /// 规则：四舍五入到分；整数部分按四位分组（个 / 万 / 亿 / 万亿）；
    /// 无角分时以“整”结尾；有角无分不写“整”；支持负数与零。
    /// </para>
    /// </summary>
    public static class ChineseAmountService
    {
        private static readonly string[] Digits = { "零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖" };

        private static readonly string[] UnitsInGroup = { "", "拾", "佰", "仟" };

        private static readonly string[] SectionUnits = { "", "万", "亿", "万亿" };

        /// <summary>
        /// AccuX 支持的最大金额（约 9999 万亿），超过时抛出异常由调用方提示。
        /// </summary>
        public const decimal MaxAmount = 9999999999999999.99m;

        /// <summary>
        /// 将金额转换为中文大写。
        /// </summary>
        public static string ToChineseAmount(decimal amount)
        {
            if (amount > MaxAmount || amount < -MaxAmount)
            {
                throw new ArgumentOutOfRangeException(nameof(amount), "金额超出 AccuX 支持范围。");
            }

            if (amount == 0m)
            {
                return "零元整";
            }

            var negative = amount < 0m;
            var absolute = Math.Abs(amount);

            // 统一舍入到分，避免后续 decimal 精度残留。
            var rounded = Math.Round(absolute, 2, MidpointRounding.AwayFromZero);

            var totalFen = decimal.ToInt64(decimal.Round(rounded * 100m, 0, MidpointRounding.AwayFromZero));
            var yuan = totalFen / 100;
            var jiao = (int)(totalFen / 10 % 10);
            var fen = (int)(totalFen % 10);

            var builder = new StringBuilder();
            if (negative)
            {
                builder.Append("负");
            }

            builder.Append(ConvertInteger(yuan));
            builder.Append("元");

            if (jiao == 0 && fen == 0)
            {
                builder.Append("整");
            }
            else
            {
                if (jiao == 0)
                {
                    // 元与分之间需要“零”，例如 壹元零伍分。
                    builder.Append("零");
                }
                else
                {
                    builder.Append(Digits[jiao]).Append("角");
                }

                if (fen != 0)
                {
                    builder.Append(Digits[fen]).Append("分");
                }
            }

            return builder.ToString();
        }

        /// <summary>
        /// 转换非负整数部分（元）。
        /// </summary>
        internal static string ConvertInteger(long value)
        {
            if (value == 0)
            {
                return "零";
            }

            var groups = new List<int>();
            var remaining = value;
            while (remaining > 0)
            {
                groups.Add((int)(remaining % 10000));
                remaining /= 10000;
            }

            var builder = new StringBuilder();
            var pendingZero = false;

            for (var index = groups.Count - 1; index >= 0; index--)
            {
                var group = groups[index];
                if (group == 0)
                {
                    // 空分组本身不输出，但要求更高位与更低非空位之间补“零”。
                    pendingZero = true;
                    continue;
                }

                // 本组不足四位（仟位为 0）时，与高位组之间存在断档，需要补“零”。
                if (builder.Length > 0 && (pendingZero || group < 1000))
                {
                    builder.Append("零");
                }

                pendingZero = false;
                builder.Append(ConvertGroup(group));
                builder.Append(SectionUnits[index]);
            }

            return builder.ToString();
        }

        /// <summary>
        /// 转换组内 0..9999。
        /// </summary>
        internal static string ConvertGroup(int group)
        {
            var builder = new StringBuilder();
            var zeroPending = false;
            var positions = new[] { 1000, 100, 10, 1 };

            for (var index = 0; index < positions.Length; index++)
            {
                var digit = group / positions[index] % 10;
                if (digit == 0)
                {
                    if (builder.Length > 0)
                    {
                        zeroPending = true;
                    }

                    continue;
                }

                if (zeroPending)
                {
                    builder.Append("零");
                    zeroPending = false;
                }

                builder.Append(Digits[digit]);
                builder.Append(UnitsInGroup[positions.Length - 1 - index]);
            }

            return builder.ToString();
        }

        /// <summary>
        /// 尝试转换，失败时返回 false 并给出错误信息。
        /// </summary>
        public static bool TryToChineseAmount(decimal amount, out string result, out string error)
        {
            result = null;
            error = null;

            try
            {
                result = ToChineseAmount(amount);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        /// <summary>
        /// 便捷方法：从文本金额转换。
        /// </summary>
        public static bool TryParse(string text, out decimal amount)
        {
            return decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out amount);
        }
    }
}
