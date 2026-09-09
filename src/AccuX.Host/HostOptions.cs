namespace AccuX.Host
{
    /// <summary>
    /// 宿主适配层配置。阈值最终默认值需通过 Excel/WPS 实测 benchmark 确定（规格 §16）。
    /// </summary>
    public sealed class HostOptions
    {
        /// <summary>超过该单元格数时提示用户确认后执行。</summary>
        public long LargeSelectionWarning { get; set; } = 100000;

        /// <summary>超过该单元格数时 V1 直接拒绝处理。</summary>
        public long MaxProcessCells { get; set; } = 500000;
    }
}
