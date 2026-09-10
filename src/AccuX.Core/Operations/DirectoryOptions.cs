using System;

namespace AccuX.Core.Operations
{
    /// <summary>
    /// “生成目录”工作表的文案与外观参数（规格 §5.2：业务决定与宿主机制的边界）。
    /// <para>
    /// “目录长什么样”（标题、表头、配色、列宽、行高、字号）属于业务决定，由业务模块提供；
    /// Host 只负责把这些参数落到 Excel/WPS COM 对象上，不再内置任何目录专属常量。
    /// </para>
    /// <para>
    /// 颜色统一使用 #RRGGBB 文本格式，避免 Core 出现宿主特有的 OLE 颜色表示。
    /// </para>
    /// </summary>
    public sealed class DirectoryOptions
    {
        /// <summary>目录工作表名称，例如“目录”。</summary>
        public string WorksheetName { get; set; } = string.Empty;

        /// <summary>标题后缀：标题 = 工作簿文件名 + 后缀。</summary>
        public string TitleSuffix { get; set; } = string.Empty;

        /// <summary>工作簿文件名为空时使用的标题。</summary>
        public string FallbackTitle { get; set; } = string.Empty;

        /// <summary>表头文字，固定三列。</summary>
        public string[] Headers { get; set; } = Array.Empty<string>();

        /// <summary>三列列宽。</summary>
        public double[] ColumnWidths { get; set; } = Array.Empty<double>();

        /// <summary>标题行行高。</summary>
        public double TitleRowHeight { get; set; }

        /// <summary>表头行行高。</summary>
        public double HeaderRowHeight { get; set; }

        /// <summary>数据行行高。</summary>
        public double BodyRowHeight { get; set; }

        /// <summary>标题字号。</summary>
        public int TitleFontSize { get; set; }

        /// <summary>表头字号。</summary>
        public int HeaderFontSize { get; set; }

        /// <summary>标题背景色（#RRGGBB）。</summary>
        public string TitleBackgroundColor { get; set; } = string.Empty;

        /// <summary>表头背景色（#RRGGBB）。</summary>
        public string HeaderBackgroundColor { get; set; } = string.Empty;

        /// <summary>隔行底色（#RRGGBB）。</summary>
        public string AlternateRowColor { get; set; } = string.Empty;

        /// <summary>边框颜色（#RRGGBB）。</summary>
        public string BorderColor { get; set; } = string.Empty;

        /// <summary>标题文字颜色（#RRGGBB）。</summary>
        public string TitleTextColor { get; set; } = string.Empty;

        /// <summary>正文文字颜色（#RRGGBB）。</summary>
        public string BodyTextColor { get; set; } = string.Empty;

        /// <summary>名称列超链接文字颜色（#RRGGBB）。</summary>
        public string HyperlinkColor { get; set; } = string.Empty;
    }
}
