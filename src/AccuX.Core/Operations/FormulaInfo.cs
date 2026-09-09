namespace AccuX.Core.Operations
{
    /// <summary>
    /// 公式种类（规格 §5.2）。
    /// 由 Host 识别，Core / Module 只依据该枚举决定是否转换。
    /// </summary>
    public enum FormulaKind
    {
        /// <summary>普通公式。</summary>
        Normal,

        /// <summary>数组公式（CSE 公式）。</summary>
        Array,

        /// <summary>动态数组公式。</summary>
        DynamicArray,

        /// <summary>无法稳定识别或无法安全写回的公式。</summary>
        Unsupported
    }

    /// <summary>
    /// Host 提供给 Core / Module 的规范化公式信息。
    /// <para>
    /// <see cref="Expression"/> 是规范化后的公式表达式（统一以 '=' 开头、参数分隔符与数字 literal 已规范化）；
    /// 业务层不得区分 Excel Formula / Formula2 / WPS Formula 或本地区域格式 API。
    /// </para>
    /// </summary>
    public sealed class FormulaInfo
    {
        public FormulaInfo(string expression, FormulaKind kind, bool canTransform)
        {
            Expression = expression;
            Kind = kind;
            CanTransform = canTransform;
        }

        /// <summary>规范化公式表达式，统一以 '=' 开头。</summary>
        public string Expression { get; }

        /// <summary>公式种类。</summary>
        public FormulaKind Kind { get; }

        /// <summary>是否允许安全修改；为 false 时默认跳过。</summary>
        public bool CanTransform { get; }

        public override string ToString()
        {
            return Kind + ":" + Expression;
        }
    }
}
