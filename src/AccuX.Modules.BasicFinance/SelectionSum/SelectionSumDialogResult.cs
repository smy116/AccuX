namespace AccuX.Modules.BasicFinance.SelectionSum
{
    /// <summary>
    /// 选区求和结果对话框中用户执行的复制动作。
    /// </summary>
    public enum SelectionSumCopyKind
    {
        None = 0,
        Amount = 1,
        WanAmount = 2,
        ChineseAmount = 3
    }

    /// <summary>
    /// 选区求和结果对话框的交互结果。
    /// </summary>
    public sealed class SelectionSumDialogResult
    {
        private SelectionSumDialogResult(
            bool wasCancelled,
            SelectionSumCopyKind copyKind,
            bool copySucceeded)
        {
            WasCancelled = wasCancelled;
            CopyKind = copyKind;
            CopySucceeded = copySucceeded;
        }

        /// <summary>用户是否直接关闭了对话框。</summary>
        public bool WasCancelled { get; }

        /// <summary>用户点击的复制类型。</summary>
        public SelectionSumCopyKind CopyKind { get; }

        /// <summary>剪切板复制是否成功。</summary>
        public bool CopySucceeded { get; }

        public static SelectionSumDialogResult Cancelled()
        {
            return new SelectionSumDialogResult(true, SelectionSumCopyKind.None, false);
        }

        public static SelectionSumDialogResult CopyAttempted(
            SelectionSumCopyKind copyKind,
            bool copySucceeded)
        {
            return new SelectionSumDialogResult(false, copyKind, copySucceeded);
        }
    }
}
