using System;
using System.Windows;
using AccuX.Modules.BasicFinance.UI;

namespace AccuX.Modules.BasicFinance.SelectionSum
{
    /// <summary>
    /// 选区求和多格式复制窗口。
    /// 窗口只负责显示结果和转发用户的复制动作，不实现财务计算。
    /// </summary>
    public partial class SelectionSumWindow : Window
    {
        private readonly Func<string, bool> _tryCopy;

        public SelectionSumWindow(
            SelectionSumResult result,
            IntPtr ownerHandle,
            Func<string, bool> tryCopy)
        {
            if (result == null)
            {
                throw new ArgumentNullException(nameof(result));
            }

            InitializeComponent();

            _tryCopy = tryCopy;
            WindowHelper.AttachOwner(this, ownerHandle);

            AmountBox.Text = result.FormattedTotal;
            WanAmountBox.Text = result.FormattedWanTotal;
            ChineseAmountBox.Text = result.ChineseTotal;
        }

        public SelectionSumDialogResult Result { get; private set; }

        private void OnCopyAmount(object sender, RoutedEventArgs e)
        {
            CompleteCopy(SelectionSumCopyKind.Amount, AmountBox.Text);
        }

        private void OnCopyWanAmount(object sender, RoutedEventArgs e)
        {
            CompleteCopy(SelectionSumCopyKind.WanAmount, WanAmountBox.Text);
        }

        private void OnCopyChineseAmount(object sender, RoutedEventArgs e)
        {
            CompleteCopy(SelectionSumCopyKind.ChineseAmount, ChineseAmountBox.Text);
        }

        private void CompleteCopy(SelectionSumCopyKind copyKind, string text)
        {
            var copied = false;
            try
            {
                copied = _tryCopy != null && _tryCopy(text);
            }
            catch
            {
                // 复制失败时仍然关闭窗口；调用方会通过结果和日志记录失败。
                copied = false;
            }

            Result = SelectionSumDialogResult.CopyAttempted(copyKind, copied);
            DialogResult = true;
        }
    }
}
