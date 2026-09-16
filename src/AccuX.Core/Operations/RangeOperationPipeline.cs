using System;
using System.Collections.Generic;
using System.Diagnostics;
using AccuX.Core.Cells;
using AccuX.Core.Logging;

namespace AccuX.Core.Operations
{
    /// <summary>
    /// 统一 Range 操作管线（规格 §10 / §16）。
    /// <para>
    /// 只负责通用编排，通过 <see cref="IRangeOperationHost"/> 与宿主交互：
    /// 验证 → 批量读取 → 单元格分类 → 调用业务 transform → 生成完整 WritePlan →
    /// 写入前最终检查 → 批量写回 → 恢复宿主状态 → 统计与日志。
    /// </para>
    /// <para>
    /// 不引用 AccuX.Host、不引用 Excel/WPS Interop、不持有 COM 对象、不判断 HostKind 分支。
    /// 遵循 fail before write：写回前完成全部读取、分类、计算与待写结果生成。
    /// </para>
    /// </summary>
    public sealed class RangeOperationPipeline
    {
        private readonly IRangeOperationHost _host;
        private readonly ILogger _logger;
        private readonly IReadOnlyList<IRangeTransformPostProcessor> _postProcessors;

        public RangeOperationPipeline(IRangeOperationHost host, ILogger logger, IEnumerable<IRangeTransformPostProcessor> postProcessors = null)
        {
            _host = host ?? throw new ArgumentNullException(nameof(host));
            _logger = logger ?? NullLogger.Instance;
            _postProcessors = postProcessors == null
                ? Array.Empty<IRangeTransformPostProcessor>()
                : new List<IRangeTransformPostProcessor>(postProcessors);
        }

        /// <summary>
        /// 从当前 Selection 捕获唯一 RangeTarget（规格 §10 / §19）。
        /// 一次 Command 只允许调用一次；Pipeline 只是转发给宿主，不缓存任何 COM 对象。
        /// </summary>
        public RangeTarget CaptureTarget()
        {
            return _host.CaptureTarget();
        }

        /// <summary>
        /// 按已捕获的 RangeTarget 批量读取选区数据，但不执行任何写回。
        /// 供只读业务命令复用 Host 的统一读取、分类与隐藏行列识别能力。
        /// </summary>
        public RangeReadResult Read(RangeTarget target)
        {
            if (target == null)
            {
                throw new ArgumentNullException(nameof(target));
            }

            return _host.Read(target);
        }

        /// <summary>
        /// 执行一次完整的 Range 操作。
        /// </summary>
        /// <param name="transform">业务转换。</param>
        /// <param name="context">操作上下文（含已捕获的 RangeTarget 与参数）。</param>
        /// <param name="stateOptions">宿主状态选项；null 表示不修改宿主全局状态。</param>
        public OperationResult Execute(IRangeTransform transform, OperationContext context, HostStateOptions stateOptions)
        {
            if (transform == null)
            {
                throw new ArgumentNullException(nameof(transform));
            }

            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            var target = context.Target;
            var stopwatch = Stopwatch.StartNew();
            var stats = new TransformStats();

            try
            {
                if (target == null)
                {
                    throw new ArgumentException("操作目标不能为空。", nameof(context));
                }

                // AccuX 不对合并区域执行批量写回。Host 在捕获和重新校验阶段都会
                // 设置该标志；Pipeline 再做一次纯 CLR 层拦截，避免其他 Host
                // 实现忘记处理时仍然修改合并单元格。
                if (target.ContainsMergedCells)
                {
                    stopwatch.Stop();
                    return new OperationResult(target, stats)
                    {
                        Success = false,
                        Duration = stopwatch.Elapsed,
                        Message = "选区包含合并单元格，当前操作不支持。",
                        Issues = new[] { "选区包含合并单元格，当前操作不支持。" }
                    };
                }

                // 1. 批量读取（唯一一次读取，之后不再依赖当前 Selection）。
                var readResult = _host.Read(target);
                stats.TotalCells = readResult.Cells.Count;
                stats.ContainsHiddenRows = readResult.ContainsHiddenRows;
                stats.ContainsHiddenColumns = readResult.ContainsHiddenColumns;

                // 2. 生成完整 WritePlan（内存计算，尚未写入任何数据）。
                var writePlan = new RangeWritePlan(target);

                foreach (var cell in readResult.Cells)
                {
                    // 隐藏状态由 Host 在批量读取阶段识别，所有业务功能在公共管线统一跳过。
                    // 这样业务 Transform 无需各自接触 Excel/WPS COM 或重复实现可见性规则。
                    if (cell != null && cell.IsHidden)
                    {
                        stats.RecordSkip(SkipReason.Hidden);
                        continue;
                    }

                    var outcome = transform.Transform(cell, context);
                    if (outcome == null || !outcome.HasWrite)
                    {
                        stats.RecordSkip(outcome?.SkipReason ?? SkipReason.None);
                        continue;
                    }

                    if (outcome.IsFormulaWrite)
                    {
                        writePlan.AddFormula(cell.Row, cell.Column, outcome.Formula);
                        stats.FormulaWrites++;
                    }
                    else
                    {
                        writePlan.AddValue(cell.Row, cell.Column, outcome.Value, outcome.NumberFormat);
                        stats.ValueWrites++;
                    }

                    stats.ProcessedCells++;
                }

                // 3. 业务后处理（例如整体调整 NumberFormat）。
                foreach (var postProcessor in _postProcessors)
                {
                    postProcessor.Process(context, readResult, writePlan);
                }

                if (writePlan.IsEmpty)
                {
                    stopwatch.Stop();
                    return new OperationResult(target, stats)
                    {
                        Success = true,
                        Duration = stopwatch.Elapsed,
                        Message = stats.GetSkipCount(SkipReason.Hidden) > 0
                            ? BuildMessage(stats)
                            : "选区中没有可处理的单元格。"
                    };
                }

                // 4. 写入前最终检查。
                var check = _host.ValidateWrite(target, writePlan);
                if (check == null || !check.CanWrite)
                {
                    stopwatch.Stop();
                    return OperationResult.Failed(target, ToArray(check));
                }

                // 5. 写回（仅在需要时才创建宿主状态作用域）。
                IHostStateScope scope = null;
                try
                {
                    if (stateOptions != null)
                    {
                        scope = _host.BeginStateScope(stateOptions);
                    }

                    _host.Write(target, writePlan);
                }
                finally
                {
                    scope?.Dispose();
                }

                stopwatch.Stop();
                var result = new OperationResult(target, stats)
                {
                    Success = true,
                    Duration = stopwatch.Elapsed,
                    Message = BuildMessage(stats)
                };

                LogOperation(transform, context, result);
                return result;
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                _logger.Error("RangeOperationPipeline 执行失败: " + transform.Id, ex);

                return new OperationResult(target, stats)
                {
                    Success = false,
                    Duration = stopwatch.Elapsed,
                    Message = ex.Message,
                    Issues = new[] { ex.Message }
                };
            }
        }

        private void LogOperation(IRangeTransform transform, OperationContext context, OperationResult result)
        {
            _logger.LogOperation(new OperationLogEntry
            {
                CommandId = context.CommandId,
                ModuleId = context.ModuleId,
                HostKind = _host.Context?.HostKind.ToString(),
                HostVersion = _host.Context?.HostVersion,
                WorksheetName = context.Target.WorksheetName,
                RangeAddress = context.Target.Address,
                CellCount = context.Target.CellCount,
                Duration = result.Duration,
                Result = "processed=" + result.Stats.ProcessedCells
                    + ",skipped=" + result.Stats.SkippedCells
                    + ",hidden=" + result.Stats.GetSkipCount(SkipReason.Hidden)
            });
        }

        private static string[] ToArray(WriteCheckResult check)
        {
            if (check == null || check.Issues.Count == 0)
            {
                return new[] { "目标区域当前不可写入。" };
            }

            var array = new string[check.Issues.Count];
            for (var i = 0; i < check.Issues.Count; i++)
            {
                array[i] = check.Issues[i];
            }

            return array;
        }

        private static string BuildMessage(TransformStats stats)
        {
            var hiddenCount = stats.GetSkipCount(SkipReason.Hidden);
            var hiddenDimension = stats.ContainsHiddenRows && stats.ContainsHiddenColumns
                ? "行/列"
                : stats.ContainsHiddenRows
                    ? "行"
                    : "列";

            return "成功处理：" + stats.ProcessedCells
                + "，跳过文本：" + stats.GetSkipCount(SkipReason.Text)
                + "，跳过空白：" + stats.GetSkipCount(SkipReason.Blank)
                + "，跳过日期：" + stats.GetSkipCount(SkipReason.Date)
                + "，跳过逻辑值：" + stats.GetSkipCount(SkipReason.Boolean)
                + "，跳过错误：" + stats.GetSkipCount(SkipReason.Error)
                + "，跳过公式：" + stats.GetSkipCount(SkipReason.FormulaNotTransformable)
                + (hiddenCount > 0 ? "，跳过隐藏" + hiddenDimension + "：" + hiddenCount : string.Empty);
        }
    }

    /// <summary>
    /// 业务后处理扩展点，在 WritePlan 生成后、写入前调用。
    /// 只处理 CLR 数据，不得访问宿主 COM 对象。
    /// </summary>
    public interface IRangeTransformPostProcessor
    {
        void Process(OperationContext context, RangeReadResult readResult, RangeWritePlan writePlan);
    }
}
