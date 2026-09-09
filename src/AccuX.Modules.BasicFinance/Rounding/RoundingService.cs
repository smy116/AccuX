using System;

namespace AccuX.Modules.BasicFinance.Rounding
{
    /// <summary>
    /// 一键舍入的纯 C# 业务算法（规格 §12.2）。
    /// 使用 <see cref="MidpointRounding.AwayFromZero"/>，不依赖 Math.Round 默认中点行为。
    /// </summary>
    public static class RoundingService
    {
        /// <summary>
        /// 按指定小数位进行财务舍入（四舍五入，远离零）。
        /// </summary>
        public static decimal Round(decimal value, int digits)
        {
            if (digits < 0 || digits > RoundingOptions.MaxDigits)
            {
                throw new ArgumentOutOfRangeException(nameof(digits), digits, "小数位数超出允许范围。");
            }

            return Math.Round(value, digits, MidpointRounding.AwayFromZero);
        }

        /// <summary>
        /// 判断舍入是否真的改变了数值，避免无意义的写入。
        /// </summary>
        public static bool WouldChange(decimal value, int digits)
        {
            return Round(value, digits) != value;
        }
    }
}
