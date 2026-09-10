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
        /// 判断当前活动工作簿中是否已存在“目录”工作表。
        /// </summary>
        /// <returns>已存在返回 true。</returns>
        /// <exception cref="HostOperationException">工作簿不可用时抛出。</exception>
        bool DirectoryWorksheetExists();

        /// <summary>
        /// 在当前活动工作簿的最前面生成目录工作表。
        /// </summary>
        /// <param name="replaceExisting">
        /// 为 true 时先删除已存在的“目录”工作表再重新生成；
        /// 为 false 时若已存在“目录”工作表则抛出异常。
        /// </param>
        /// <returns>目录中包含的可见工作表数量。</returns>
        /// <exception cref="HostOperationException">工作簿不可用或目录无法生成时抛出。</exception>
        int GenerateDirectory(bool replaceExisting);
    }
}
