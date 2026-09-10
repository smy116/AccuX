namespace AccuX.Modules.Compare
{
    public sealed class CompareConfig
    {
        public long LargeSelectionWarning { get; set; } = 100000;

        public long MaxProcessCells { get; set; } = 500000;

        public string FirstOnlyColor { get; set; } = "#FFFF66";

        public string SecondOnlyColor { get; set; } = "#FF9999";

        public string SameColor { get; set; } = "#CCFFCC";
    }
}
