namespace AccuX.Core.Operations
{
    /// <summary>
    /// 工作簿目录操作的宿主边界。
    /// <para>
    /// 目录生成需要访问 Workbook / Worksheet COM 对象，因此由 Host 实现，
    /// 业务模块只通过本接口触发操作并接收生成数量；文案与外观由
    /// <see cref="DirectoryOptions"/> 描述，见该类型的说明。
    /// </para>
    /// </summary>
    public interface IWorkbookDirectoryHost
    {
        /// <summary>
        /// 判断当前活动工作簿中是否已存在指定名称的工作表。
        /// </summary>
        /// <param name="worksheetName">工作表名称，例如“目录”。</param>
        /// <returns>已存在返回 true。</returns>
        /// <exception cref="HostOperationException">工作簿不可用时抛出。</exception>
        bool DirectoryWorksheetExists(string worksheetName);

        /// <summary>
        /// 在当前活动工作簿的最前面生成目录工作表。
        /// </summary>
        /// <param name="options">目录的文案与外观参数，由业务模块提供。</param>
        /// <param name="replaceExisting">
        /// 为 true 时先删除已存在的同名工作表再重新生成；
        /// 为 false 时若已存在同名工作表则抛出异常。
        /// </param>
        /// <returns>目录中包含的可见工作表数量。</returns>
        /// <exception cref="HostOperationException">工作簿不可用或目录无法生成时抛出。</exception>
        int GenerateDirectory(DirectoryOptions options, bool replaceExisting);
    }
}
