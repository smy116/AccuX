using System;
using AccuX.Modules.BasicFinance.Rounding;

namespace AccuX.Modules.BasicFinance.AmountConversion
{
    /// <summary>
    /// 金额折合的纯 C# 业务算法（规格 §13.2）。
    /// 核心计算使用 decimal；折合后按指定小数位舍入。
    /// </summary>
    public static class AmountConversionService
    {
        /// <summary>
        /// 对单个金额执行折合。除法模式下 Rate 为 0 抛出异常（调用方应已在 UI 层阻止）。
        /// </summary>
        public static decimal Convert(decimal amount, AmountConversionOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (options.Mode == ConversionMode.Divide && options.Rate == 0m)
            {
                throw new DivideByZeroException("折合率不能为 0。");
            }

            var converted = options.Mode == ConversionMode.Divide
                ? amount / options.Rate
                : amount * options.Rate;

            return RoundingService.Round(converted, options.Digits);
        }

        /// <summary>
        /// 对单个金额执行折合并按需生成带“万”字的文本结果。
        /// </summary>
        public static object ConvertToResult(decimal amount, AmountConversionOptions options)
        {
            var value = Convert(amount, options);
            if (options.AppendWanSuffix)
            {
                return value.ToString("0." + new string('0', Math.Max(0, options.Digits))) + "万";
            }

            return value;
        }

        /// <summary>
        /// 生成公式变换后的规范化表达式：原公式 × 或 ÷ Rate，必要时套 ROUND。
        /// </summary>
        public static string BuildFormula(string normalizedExpression, AmountConversionOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (options.Mode == ConversionMode.Divide && options.Rate == 0m)
            {
                throw new DivideByZeroException("折合率不能为 0。");
            }

            var operatorSymbol = options.Mode == ConversionMode.Divide ? "/" : "*";
            var arithmetic = Common.FormulaTransformService.ApplyArithmetic(
                normalizedExpression,
                operatorSymbol,
                options.Rate);

            return Common.FormulaTransformService.WrapFunction(
                arithmetic,
                "ROUND",
                options.Digits.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
    }
}
