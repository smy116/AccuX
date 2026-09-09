using System;
using System.Collections.Generic;

namespace AccuX.Core.Operations
{
    /// <summary>
    /// 处理结果统计（规格 §12.4）。
    /// </summary>
    public sealed class TransformStats
    {
        private readonly Dictionary<SkipReason, int> _skipCounts = new Dictionary<SkipReason, int>();

        public int TotalCells { get; internal set; }

        public int ProcessedCells { get; internal set; }

        public int ValueWrites { get; internal set; }

        public int FormulaWrites { get; internal set; }

        /// <summary>本次操作的目标选区是否包含被隐藏的行。</summary>
        public bool ContainsHiddenRows { get; internal set; }

        /// <summary>本次操作的目标选区是否包含被隐藏的列。</summary>
        public bool ContainsHiddenColumns { get; internal set; }

        public int SkippedCells
        {
            get { return TotalCells - ProcessedCells; }
        }

        internal void RecordSkip(SkipReason reason)
        {
            _skipCounts.TryGetValue(reason, out var count);
            _skipCounts[reason] = count + 1;
        }

        public int GetSkipCount(SkipReason reason)
        {
            return _skipCounts.TryGetValue(reason, out var count) ? count : 0;
        }

        public IReadOnlyDictionary<SkipReason, int> SkipCounts
        {
            get { return _skipCounts; }
        }
    }

    /// <summary>
    /// 一次 Range 操作的完整结果，供 Command 层生成用户提示。
    /// </summary>
    public sealed class OperationResult
    {
        public OperationResult(RangeTarget target, TransformStats stats)
        {
            Target = target;
            Stats = stats ?? new TransformStats();
        }

        public RangeTarget Target { get; }

        public TransformStats Stats { get; }

        public bool Success { get; internal set; }

        /// <summary>面向用户的简短说明。</summary>
        public string Message { get; internal set; }

        /// <summary>失败原因（用户可读）。</summary>
        public IReadOnlyList<string> Issues { get; internal set; } = Array.Empty<string>();

        public TimeSpan Duration { get; internal set; }

        public static OperationResult Failed(RangeTarget target, params string[] issues)
        {
            return new OperationResult(target, new TransformStats())
            {
                Success = false,
                Issues = issues ?? Array.Empty<string>(),
                Message = issues != null && issues.Length > 0 ? issues[0] : "操作失败。"
            };
        }
    }
}
