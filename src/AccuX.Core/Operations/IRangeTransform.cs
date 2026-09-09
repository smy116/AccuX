using System.Collections.Generic;

namespace AccuX.Core.Operations
{
    /// <summary>
    /// 业务转换的跳过原因分类，用于结果统计与提示（规格 §12.4 / §15）。
    /// </summary>
    public enum SkipReason
    {
        None = 0,
        Blank,
        Text,
        Date,
        Boolean,
        Error,
        UnsupportedType,
        FormulaNotTransformable,
        MergeCell,
        NotWritable,
        Hidden
    }

    /// <summary>
    /// 单个单元格的转换决策结果。
    /// </summary>
    public sealed class CellTransformOutcome
    {
        private CellTransformOutcome(CellData cell)
        {
            Cell = cell;
        }

        public CellData Cell { get; }

        /// <summary>是否产生写入。</summary>
        public bool HasWrite { get; private set; }

        public SkipReason SkipReason { get; private set; } = SkipReason.None;

        public object Value { get; private set; }

        public string Formula { get; private set; }

        public string NumberFormat { get; private set; }

        public bool IsFormulaWrite
        {
            get { return Formula != null; }
        }

        public static CellTransformOutcome Skip(CellData cell, SkipReason reason)
        {
            return new CellTransformOutcome(cell)
            {
                SkipReason = reason
            };
        }

        public static CellTransformOutcome WriteValue(CellData cell, object value, string numberFormat = null)
        {
            return new CellTransformOutcome(cell)
            {
                HasWrite = true,
                Value = value,
                NumberFormat = numberFormat
            };
        }

        public static CellTransformOutcome WriteFormula(CellData cell, string formula)
        {
            return new CellTransformOutcome(cell)
            {
                HasWrite = true,
                Formula = formula
            };
        }
    }

    /// <summary>
    /// 业务模块实现的纯 CLR 转换契约。
    /// <para>
    /// 业务功能只提供“怎么转换”的逻辑：不遍历 COM Cell、不处理 Excel/WPS 差异、不读写宿主状态。
    /// </para>
    /// </summary>
    public interface IRangeTransform
    {
        /// <summary>转换标识，用于日志与统计。</summary>
        string Id { get; }

        /// <summary>
        /// 对单个单元格给出转换决策。
        /// 返回 null 视为跳过且不计数；通常应显式返回 <see cref="CellTransformOutcome.Skip"/>。
        /// </summary>
        CellTransformOutcome Transform(CellData cell, OperationContext context);

        /// <summary>
        /// 转换是否会产生公式写入；用于 Pipeline 提前决定是否需要禁用重算等宿主状态。
        /// </summary>
        bool MayWriteFormula { get; }
    }

    /// <summary>
    /// 转换执行上下文。只含 CLR 数据，禁止携带 COM 引用。
    /// </summary>
    public sealed class OperationContext
    {
        public OperationContext(string commandId, string moduleId, RangeTarget target, IDictionary<string, object> parameters = null)
        {
            CommandId = commandId;
            ModuleId = moduleId;
            Target = target;
            Parameters = parameters ?? new Dictionary<string, object>();
        }

        public string CommandId { get; }

        public string ModuleId { get; }

        public RangeTarget Target { get; }

        public IDictionary<string, object> Parameters { get; }

        public T GetParameter<T>(string name, T defaultValue = default)
        {
            if (Parameters != null && Parameters.TryGetValue(name, out var value))
            {
                if (value is T typed)
                {
                    return typed;
                }

                if (value != null)
                {
                    try
                    {
                        return (T)System.Convert.ChangeType(value, typeof(T));
                    }
                    catch
                    {
                        return defaultValue;
                    }
                }
            }

            return defaultValue;
        }
    }
}
