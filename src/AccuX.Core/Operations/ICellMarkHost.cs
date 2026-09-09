using System;

namespace AccuX.Core.Operations
{
    /// <summary>
    /// 当前选区标记操作的宿主边界。
    /// <para>
    /// Host 负责 RangeTarget 的解析、可见行列判断、保护状态检查以及底色写入；
    /// 业务模块只传递经过固定的选区和 #RRGGBB 颜色值，不接触 Excel/WPS COM。
    /// </para>
    /// </summary>
    public interface ICellMarkHost
    {
        /// <summary>
        /// 将指定颜色应用到目标选区中的可见单元格，并返回实际修改的单元格数量。
        /// </summary>
        /// <param name="target">Command 开始时固定的选区目标。</param>
        /// <param name="hexColor">HTML 形式的颜色值，例如 #18be6a。</param>
        /// <exception cref="HostOperationException">目标不可用、不可写或颜色无效时抛出。</exception>
        long ApplyVisibleBackgroundColor(RangeTarget target, string hexColor);
    }
}
