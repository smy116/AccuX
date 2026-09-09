namespace AccuX.Core.Operations
{
    /// <summary>
    /// 传统单元格批注宿主边界。
    /// <para>
    /// Host 负责选择解析、工作簿/工作表身份校验、保护状态检查以及 COM 批注 API 调用；
    /// 业务模块只处理 CLR 文本和批注类型。
    /// </para>
    /// </summary>
    public interface ICellCommentHost
    {
        /// <summary>
        /// 从当前 Selection 固定一个批注目标。连续选区取左上角；多区域按行、列取最小单元格。
        /// </summary>
        CellCommentTarget CaptureCommentTarget();

        /// <summary>读取已固定目标上的传统批注。</summary>
        CellComment ReadComment(CellCommentTarget target);

        /// <summary>覆盖或新增已固定目标上的传统批注。</summary>
        void SaveComment(CellCommentTarget target, string text);

        /// <summary>删除已固定目标上的传统批注；目标没有批注时视为成功。</summary>
        void DeleteComment(CellCommentTarget target);
    }
}
