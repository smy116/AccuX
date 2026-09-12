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

        // 保留旧版字段，使旧配置模型和回滚到旧版本的代码仍可读取；新代码统一从 settings 读取。
        public long LargeSelectionWarning { get; set; } = 100000;

        public long MaxProcessCells { get; set; } = 500000;
    }
}
