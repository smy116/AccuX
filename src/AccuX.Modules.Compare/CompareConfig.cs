namespace AccuX.Modules.Compare
{
    public sealed class CompareConfig
    {
        // 保留旧版字段，使旧配置模型和回滚到旧版本的代码仍可读取；新代码统一从 settings 读取。
        public long LargeSelectionWarning { get; set; } = 100000;

        public long MaxProcessCells { get; set; } = 500000;

        public string FirstOnlyColor { get; set; } = "#FFFF66";

        public string SecondOnlyColor { get; set; } = "#FFFF66";

        public string SameColor { get; set; } = "#CCFFCC";
    }
}
