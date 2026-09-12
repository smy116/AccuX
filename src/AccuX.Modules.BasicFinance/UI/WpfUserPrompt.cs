using System;
using System.Windows;
using AccuX.Core.Operations;
using AccuX.Modules.BasicFinance.AmountConversion;
using AccuX.Modules.BasicFinance.Common;
using AccuX.Modules.BasicFinance.Comment;
using AccuX.Modules.BasicFinance.Rounding;
using AccuX.Modules.BasicFinance.SelectionSum;
using AccuX.Modules.BasicFinance.UI;

namespace AccuX.Modules.BasicFinance.UI
{
    /// <summary>
    /// 基于 WPF 的 <see cref="IUserPrompt"/> 实现。
    /// 只负责界面交互；所有财务计算由业务 Service 完成（规格 §20：WPF 不实现财务计算逻辑）。
    /// </summary>
    public sealed class WpfUserPrompt : IUserPrompt, ICommentPromptSession
    {
        private readonly IHostContext _host;
        private long _largeSelectionWarning;
        private readonly int _defaultRoundDigits;

        public WpfUserPrompt(IHostContext host, long largeSelectionWarning, int defaultRoundDigits = 2)
        {
            _host = host;
            _largeSelectionWarning = largeSelectionWarning;
            _defaultRoundDigits = defaultRoundDigits >= 0 && defaultRoundDigits <= RoundingOptions.MaxDigits
                ? defaultRoundDigits
                : 2;
        }

        public void UpdateLargeSelectionWarning(long largeSelectionWarning)
        {
            if (largeSelectionWarning > 0)
            {
                _largeSelectionWarning = largeSelectionWarning;
            }
        }

        private IntPtr OwnerHandle
        {
            get { return _host?.MainWindowHandle ?? IntPtr.Zero; }
        }

        public RoundingOptions AskRoundingOptions(RangeTarget target)
        {
            var window = new RoundingWindow(_defaultRoundDigits, OwnerHandle);
            var accepted = window.ShowDialog();
            return accepted == true ? new RoundingOptions(window.Digits) : null;
        }

        public AmountConversionOptions AskAmountConversionOptions(RangeTarget target)
        {
            var viewModel = new AmountConversionViewModel();
            var window = new AmountConversionView(OwnerHandle, viewModel);
            var accepted = window.ShowDialog();
            return accepted == true ? window.Options : null;
        }

        public SelectionSumDialogResult ShowSelectionSumDialog(SelectionSumResult result)
        {
            if (result == null)
            {
                return SelectionSumDialogResult.Cancelled();
            }

            var window = new SelectionSumWindow(result, OwnerHandle, TryCopyToClipboard);
            var accepted = window.ShowDialog();
            return accepted == true && window.Result != null
                ? window.Result
                : SelectionSumDialogResult.Cancelled();
        }

        public CommentDialogResult AskComment(CommentContent existing)
        {
            return AskComment(existing, null);
        }

        public CommentDialogResult AskComment(
            CommentContent existing,
            Func<CommentDialogResult, CommentPromptSubmissionResult> submit)
        {
            var window = new CommentWindow(existing, OwnerHandle, submit);
            var accepted = window.ShowDialog();
            return accepted == true
                ? window.Result
                : CommentDialogResult.Cancelled();
        }

        public bool ConfirmLargeSelection(RangeTarget target)
        {
            if (target == null || target.CellCount <= _largeSelectionWarning)
            {
                return true;
            }

            var message = string.Format(
                System.Globalization.CultureInfo.CurrentCulture,
                "选区包含 {0} 个单元格，超过建议阈值 {1}。\n继续执行可能耗时较长，是否继续？",
                target.CellCount,
                _largeSelectionWarning);

            var result = MessageBox.Show(
                message,
                "AccuX",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            return result == MessageBoxResult.Yes;
        }

        public bool ConfirmReplaceDirectory(string worksheetName)
        {
            var message = string.Format(
                System.Globalization.CultureInfo.CurrentCulture,
                "工作簿中已存在“{0}”工作表，是否删除并重新生成？",
                worksheetName ?? string.Empty);

            var result = MessageBox.Show(
                message,
                "AccuX",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            return result == MessageBoxResult.Yes;
        }

        public void ShowMessage(string message)
        {
            if (!string.IsNullOrEmpty(message))
            {
                MessageBox.Show(message, "AccuX", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        public void ShowError(string message)
        {
            if (!string.IsNullOrEmpty(message))
            {
                MessageBox.Show(message, "AccuX", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        public bool TryCopyToClipboard(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            try
            {
                Clipboard.SetText(text, TextDataFormat.UnicodeText);
                return true;
            }
            catch
            {
                // 剪切板可能被其他进程占用；由 Command 显示不误导用户的失败提示。
                return false;
            }
        }
    }
}
