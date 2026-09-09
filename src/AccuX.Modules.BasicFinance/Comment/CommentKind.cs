namespace AccuX.Modules.BasicFinance.Comment
{
    /// <summary>批注存储类型。</summary>
    public enum CommentKind
    {
        Plain = 0,
        Encrypted = 1
    }

    /// <summary>批注窗口的结束动作。</summary>
    public enum CommentDialogAction
    {
        Cancel = 0,
        Save = 1,
        Delete = 2
    }
}
