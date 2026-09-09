using System;

namespace AccuX.Modules.BasicFinance.Common
{
    /// <summary>
    /// 公式转换服务（规格 §17）。
    /// <para>
    /// 只负责：接收规范化表达式、处理前导 '='、按调用方明确给出的业务变换生成新的规范化表达式。
    /// <b>不</b>负责防止重复包裹或判断用户是否已执行过同一操作（见 §17.2 / §17.3）。
    /// 不判断 Host 类型，不接触 Excel/WPS Formula API。
    /// </para>
    /// </summary>
    public static class FormulaTransformService
    {
        /// <summary>
        /// 去掉前导 '='，返回公式主体。
        /// </summary>
        public static string GetBody(string expression)
        {
            if (string.IsNullOrWhiteSpace(expression))
            {
                return string.Empty;
            }

            var text = expression.Trim();
            return text.StartsWith("=", StringComparison.Ordinal) ? text.Substring(1) : text;
        }

        /// <summary>
        /// 确保表达式以 '=' 开头。
        /// </summary>
        public static string EnsureLeadingEquals(string expression)
        {
            if (string.IsNullOrWhiteSpace(expression))
            {
                return string.Empty;
            }

            var text = expression.Trim();
            return text.StartsWith("=", StringComparison.Ordinal) ? text : "=" + text;
        }

        /// <summary>
        /// 用外层函数包裹原公式，例如 ROUND(A1,2)。
        /// </summary>
        public static string WrapFunction(string expression, string functionName, params string[] extraArguments)
        {
            if (string.IsNullOrWhiteSpace(functionName))
            {
                throw new ArgumentException("函数名不能为空。", nameof(functionName));
            }

            var body = GetBody(expression);
            if (string.IsNullOrEmpty(body))
            {
                throw new ArgumentException("公式表达式不能为空。", nameof(expression));
            }

            var arguments = body;
            if (extraArguments != null)
            {
                foreach (var argument in extraArguments)
                {
                    arguments += "," + argument;
                }
            }

            return "=" + functionName + "(" + arguments + ")";
        }

        /// <summary>
        /// 用外层括号包裹原公式，便于后续追加算术运算。
        /// </summary>
        public static string WrapParentheses(string expression)
        {
            var body = GetBody(expression);
            if (string.IsNullOrEmpty(body))
            {
                throw new ArgumentException("公式表达式不能为空。", nameof(expression));
            }

            return "=(" + body + ")";
        }

        /// <summary>
        /// 在原公式外层追加乘法/除法运算。
        /// </summary>
        public static string ApplyArithmetic(string expression, string operatorSymbol, decimal operand)
        {
            if (operatorSymbol != "*" && operatorSymbol != "/")
            {
                throw new ArgumentException("仅支持 * 或 / 运算。", nameof(operatorSymbol));
            }

            var body = GetBody(expression);
            if (string.IsNullOrEmpty(body))
            {
                throw new ArgumentException("公式表达式不能为空。", nameof(expression));
            }

            return "=(" + body + ")" + operatorSymbol + FormatLiteral(operand);
        }

        /// <summary>
        /// 将 decimal 格式化为公式中安全的数字字面量（InvariantCulture，去掉多余尾零）。
        /// </summary>
        public static string FormatLiteral(decimal value)
        {
            var text = value.ToString("0.############################", System.Globalization.CultureInfo.InvariantCulture);
            return text;
        }
    }
}
