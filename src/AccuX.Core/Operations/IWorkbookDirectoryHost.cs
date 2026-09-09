namespace AccuX.Core.Operations
{
    /// <summary>
    /// 工作簿目录操作的宿主边界。
    /// <para>
    /// 目录生成需要访问 Workbook / Worksheet COM 对象，因此由 Host 实现，
    /// 业务模块只通过本接口触发操作并接收生成数量。
    /// </para>
    /// </summary>
    public interface IWorkbookDirectoryHost
    {
        /// <summary>
        /// 在当前活动工作簿的最前面生成目录工作表。
        /// </summary>
        /// <returns>目录中包含的可见工作表数量。</returns>
        /// <exception cref="HostOperationException">工作簿不可用或目录无法生成时抛出。</exception>
        int GenerateDirectory();
    }
}
