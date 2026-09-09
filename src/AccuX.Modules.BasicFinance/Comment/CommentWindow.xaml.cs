using System;
using System.Windows;
using System.Windows.Input;
using AccuX.Modules.BasicFinance.Common;
using AccuX.Modules.BasicFinance.UI;

namespace AccuX.Modules.BasicFinance.Comment
{
    /// <summary>批注助手编辑窗口，只负责输入、计数和用户动作。</summary>
    public partial class CommentWindow : Window
    {
        private readonly CommentContent _existing;
        private readonly Func<CommentDialogResult, CommentPromptSubmissionResult> _submit;

        public CommentWindow(
            CommentContent existing,
            IntPtr ownerHandle,
            Func<CommentDialogResult, CommentPromptSubmissionResult> submit = null)
        {
            _existing = existing ?? CommentContent.Empty;
            _submit = submit;
            InitializeComponent();

            WindowHelper.AttachOwner(this, ownerHandle);
            ContentBox.Text = _existing.IsUnreadable ? string.Empty : _existing.Text;
            PlainRadio.IsChecked = !_existing.IsEncrypted;
            EncryptedRadio.IsChecked = _existing.IsEncrypted;

            if (_existing.IsUnreadable)
            {
                ContentBox.IsReadOnly = true;
                PlainRadio.IsEnabled = false;
                EncryptedRadio.IsEnabled = false;
                HintText.Text = _existing.ErrorMessage + " 可使用“删除”移除该批注。";
            }
            else if (_existing.IsEncrypted)
            {
                HintText.Text = "加密批注绑定当前 Windows 用户。";
            }
            else
            {
                HintText.Text = "最多1000字，空白内容不能保存。";
            }

            DeleteButton.IsEnabled = _existing.Exists;
            UpdateInputState();
        }

        public CommentDialogResult Result { get; private set; }

        private void OnTextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (_existing != null && !_existing.IsUnreadable && HintText != null)
            {
                HintText.Foreground = System.Windows.Media.Brushes.Gray;
            }

            UpdateInputState();
        }

        private void OnKindChanged(object sender, RoutedEventArgs e)
        {
            // WPF 载入 XAML 时可能在字段完成连接前触发 Checked；初始化阶段直接忽略事件。
            if (HintText == null || EncryptedRadio == null)
            {
                return;
            }

            if (_existing != null && _existing.IsUnreadable)
            {
                return;
            }

            HintText.Foreground = System.Windows.Media.Brushes.Gray;

            if (HintText != null && EncryptedRadio.IsChecked == true)
            {
                HintText.Text = "加密批注绑定当前 Windows 用户。";
            }
            else if (HintText != null)
            {
                HintText.Text = "最多1000字，空白内容不能保存。";
            }
        }

        private void OnSave(object sender, RoutedEventArgs e)
        {
            if (!CanSave())
            {
                return;
            }

            var result = CommentDialogResult.Save(
                ContentBox.Text,
                EncryptedRadio.IsChecked == true ? CommentKind.Encrypted : CommentKind.Plain);
            if (!TrySubmit(result))
            {
                return;
            }

            Result = result;
            DialogResult = true;
        }

        private void OnDelete(object sender, RoutedEventArgs e)
        {
            if (!_existing.Exists)
            {
                return;
            }

            var result = CommentDialogResult.Delete();
            if (!TrySubmit(result))
            {
                return;
            }

            Result = result;
            DialogResult = true;
        }

        private void OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Result = CommentDialogResult.Cancelled();
                DialogResult = false;
                e.Handled = true;
            }
        }

        private void UpdateInputState()
        {
            if (CounterText == null || ContentBox == null)
            {
                return;
            }

            var count = CommentTextValidator.CountTextElements(ContentBox.Text);
            CounterText.Text = CommentTextValidator.BuildCounter(ContentBox.Text);
            CounterText.Foreground = count > CommentTextValidator.MaxLength
                ? System.Windows.Media.Brushes.Red
                : System.Windows.Media.Brushes.Gray;
            SaveButton.IsEnabled = CanSave();
        }

        private bool CanSave()
        {
            return _existing != null
                && !_existing.IsUnreadable
                && ContentBox != null
                && !string.IsNullOrWhiteSpace(ContentBox.Text)
                && CommentTextValidator.IsWithinLimit(ContentBox.Text);
        }

        private bool TrySubmit(CommentDialogResult result)
        {
            if (_submit == null)
            {
                return true;
            }

            var submission = _submit(result);
            if (submission == null || submission.Success)
            {
                return true;
            }

            HintText.Text = submission.Message;
            HintText.Foreground = System.Windows.Media.Brushes.Red;
            return false;
        }
    }
}
