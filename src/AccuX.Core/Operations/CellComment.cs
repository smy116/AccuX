namespace AccuX.Core.Operations
{
    /// <summary>
    /// 宿主读取到的传统单元格批注（Excel 中称为备注）。
    /// </summary>
    public sealed class CellComment
    {
        public CellComment(bool exists, string text)
        {
            Exists = exists;
            Text = text ?? string.Empty;
        }

        public bool Exists { get; }

        public string Text { get; }

        public static CellComment None
        {
            get { return new CellComment(false, string.Empty); }
        }
    }
}
