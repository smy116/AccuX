using System;
using System.Collections.Generic;

namespace AccuX.Core.Operations
{
    /// <summary>
    /// <see cref="RangeOperationPipeline"/> 与宿主实现之间的窄接口（规格 §5.3）。
    /// <para>
    /// 存在的唯一目的是切断 Core → AccuX.Host 具体工程的依赖：Core 只依赖本接口，
    /// AccuX.Host 提供实现，AccuX.AddIn 在组合根注入。
    /// </para>
    /// <para>
    /// 实现方负责所有 Excel/WPS 差异、COM 生命周期与批量读写性能；Core 侧只处理 CLR 数据。
    /// </para>
    /// </summary>
    public interface IRangeOperationHost
    {
        /// <summary>
        /// 读取一次当前 Selection 并固化为 <see cref="RangeTarget"/>。
        /// 一次 Command 只允许调用一次；后续读写与验证都针对同一个 Target。
        /// </summary>
        /// <exception cref="HostOperationException">Selection 无法安全处理时抛出。</exception>
        RangeTarget CaptureTarget();

        /// <summary>
        /// 按 RangeTarget 批量读取 Value 与 Formula，返回归一化 CLR 数据。
        /// </summary>
        RangeReadResult Read(RangeTarget target);

        /// <summary>
        /// 写回前针对同一 RangeTarget 再次验证（身份、地址、保护状态、特殊区域）。
        /// </summary>
        WriteCheckResult ValidateWrite(RangeTarget target, RangeWritePlan writePlan);

        /// <summary>
        /// 按 RangeWritePlan 批量写回同一个 RangeTarget。
        /// 实现必须使用整块 COM 写入，禁止逐 Cell 写入。
        /// </summary>
        void Write(RangeTarget target, RangeWritePlan writePlan);

        /// <summary>
        /// 创建宿主 Application 状态作用域；只保存并恢复本次真正修改过的状态。
        /// </summary>
        IHostStateScope BeginStateScope(HostStateOptions options);

        /// <summary>
        /// 读取 RangeTarget 的 NumberFormat（用于金额折合等需要保持格式的场景）。
        /// </summary>
        string[,] ReadNumberFormats(RangeTarget target);

        /// <summary>
        /// 宿主上下文信息。
        /// </summary>
        IHostContext Context { get; }
    }

    /// <summary>
    /// 宿主操作失败异常。携带用户可读消息，由 Command 层转换为友好提示。
    /// </summary>
    public sealed class HostOperationException : Exception
    {
        public HostOperationException(string message) : base(message)
        {
        }

        public HostOperationException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
