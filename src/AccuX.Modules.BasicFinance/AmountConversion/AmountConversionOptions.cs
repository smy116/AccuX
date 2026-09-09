using System;
using System.Globalization;

namespace AccuX.Modules.BasicFinance.AmountConversion
{
    /// <summary>
    /// 折合方向（规格 §13）。
    /// </summary>
    public enum ConversionMode
    {
        /// <summary>除以折合率（折合）。</summary>
        Divide = 0,

        /// <summary>乘以折合率（换算）。</summary>
        Multiply = 1
    }

    /// <summary>
    /// 金额折合选项（规格 §13.1）。
    /// V1 采用：除百 / 除千 / 除万，可选是否添加“万”字；暂不接在线汇率。
    /// </summary>
    public sealed class AmountConversionOptions
    {
        /// <summary>折合方向。</summary>
        public ConversionMode Mode { get; set; } = ConversionMode.Divide;

        /// <summary>折合率；除百=100、除千=1000、除万=10000。</summary>
        public decimal Rate { get; set; } = 10000m;

        /// <summary>折合后保留小数位。</summary>
        public int Digits { get; set; } = 2;

        /// <summary>是否在结果后追加“万”字（结果写为文本）。</summary>
        public bool AppendWanSuffix { get; set; }

        public const int MaxDigits = 15;

        public bool IsValid
        {
            get
            {
                if (Digits < 0 || Digits > MaxDigits)
                {
                    return false;
                }

                if (Mode == ConversionMode.Divide && Rate == 0m)
                {
                    return false;
                }

                return true;
            }
        }

        /// <summary>
        /// 校验选项；除法模式下 Rate 为 0 必须阻止执行（规格 §13.4）。
        /// </summary>
        public static bool TryValidate(AmountConversionOptions options, out string error)
        {
            error = null;

            if (options == null)
            {
                error = "折合参数不能为空。";
                return false;
            }

            if (options.Digits < 0 || options.Digits > MaxDigits)
            {
                error = string.Format(CultureInfo.CurrentCulture, "小数位数必须在 0 到 {0} 之间。", MaxDigits);
                return false;
            }

            if (options.Mode == ConversionMode.Divide && options.Rate == 0m)
            {
                error = "折合率不能为 0，请重新输入。";
                return false;
            }

            return true;
        }
    }
}
