using AccuX.Core.Operations;
using AccuX.Modules.BasicFinance.AmountConversion;
using AccuX.Modules.BasicFinance.Rounding;
using AccuX.Modules.BasicFinance.SelectionSum;

namespace AccuX.Modules.BasicFinance.Common
{
    /// <summary>
    /// 模块的用户交互抽象。
    /// <para>
    /// 业务 Command 通过它获取参数与确认；具体 WPF 实现位于本模块 UI 层，
    /// 这样命令逻辑可在不启动 Excel/WPS 的情况下测试（返回预设值）。
    /// </para>
    /// </summary>
    public interface IUserPrompt
    {
        /// <summary>弹出小数位窗口；用户取消返回 null。</summary>
        RoundingOptions AskRoundingOptions(RangeTarget target);

        /// <summary>弹出金额折合窗口；用户取消返回 null。</summary>
        AmountConversionOptions AskAmountConversionOptions(RangeTarget target);

        /// <summary>显示选区求和的多格式复制窗口。</summary>
        SelectionSumDialogResult ShowSelectionSumDialog(SelectionSumResult result);

        /// <summary>
        /// 选区超过 largeSelectionWarning 时确认是否继续。
        /// 未超过阈值时直接返回 true。
        /// </summary>
        bool ConfirmLargeSelection(RangeTarget target);

        /// <summary>确认是否删除已存在的同名“目录”工作表并重新生成。</summary>
        bool ConfirmReplaceDirectory(string worksheetName);

        /// <summary>显示结果或错误提示。</summary>
        void ShowMessage(string message);

        /// <summary>显示错误提示。</summary>
        void ShowError(string message);

        /// <summary>尝试将文本复制到系统剪切板；失败返回 false。</summary>
        bool TryCopyToClipboard(string text);
    }
}
