namespace AccuX.Modules.BasicFinance
{
    /// <summary>
    /// BasicFinance 模块配置节（规格 §21）。
    /// 对应 config.json 中的 "basicFinance" 节点。
    /// </summary>
    public sealed class BasicFinanceConfig
    {
        /// <summary>一键舍入默认小数位（窗口打开时的预填值）。</summary>
        public int RoundDigits { get; set; } = 2;

        /// <summary>超过该单元格数时提示用户确认。</summary>
        public long LargeSelectionWarning { get; set; } = 100000;

        /// <summary>超过该单元格数时直接拒绝处理。</summary>
        public long MaxProcessCells { get; set; } = 500000;
    }
}
