using AccuX.Core.Operations;

namespace AccuX.Modules.BasicFinance.Directory
{
    /// <summary>
    /// “生成目录”的文案与外观定义。
    /// <para>
    /// 产品决定（标题写法、表头文字、配色、列宽、行高、字号）集中在此，
    /// 通过 <see cref="DirectoryOptions"/> 传给宿主；Host 只负责把这些参数落到 COM 上。
    /// </para>
    /// </summary>
    internal static class DirectoryTemplate
    {
        public const string WorksheetName = "目录";

        private const string TitleSuffix = "目录";
        private const string FallbackTitle = "工作簿目录";

        private const double TitleRowHeight = 30d;
        private const double HeaderRowHeight = 23d;
        private const double BodyRowHeight = 20d;
        private const int TitleFontSize = 16;
        private const int HeaderFontSize = 11;

        private const string TitleBackgroundColor = "#143146";
        private const string HeaderBackgroundColor = "#2B6391";
        private const string AlternateRowColor = "#F2F6FA";
        private const string BorderColor = "#D2DAE2";
        private const string TitleTextColor = "#FFFFFF";
        private const string BodyTextColor = "#1F2937";
        private const string HyperlinkColor = "#0563C1";

        private static readonly string[] Headers = { "序号", "名称", "备注" };
        private static readonly double[] ColumnWidths = { 8d, 30d, 42d };

        /// <summary>创建一份目录参数快照，交由宿主渲染。</summary>
        public static DirectoryOptions CreateOptions()
        {
            return new DirectoryOptions
            {
                WorksheetName = WorksheetName,
                TitleSuffix = TitleSuffix,
                FallbackTitle = FallbackTitle,
                Headers = Headers,
                ColumnWidths = ColumnWidths,
                TitleRowHeight = TitleRowHeight,
                HeaderRowHeight = HeaderRowHeight,
                BodyRowHeight = BodyRowHeight,
                TitleFontSize = TitleFontSize,
                HeaderFontSize = HeaderFontSize,
                TitleBackgroundColor = TitleBackgroundColor,
                HeaderBackgroundColor = HeaderBackgroundColor,
                AlternateRowColor = AlternateRowColor,
                BorderColor = BorderColor,
                TitleTextColor = TitleTextColor,
                BodyTextColor = BodyTextColor,
                HyperlinkColor = HyperlinkColor
            };
        }
    }
}
