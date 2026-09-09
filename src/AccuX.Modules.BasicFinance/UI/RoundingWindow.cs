using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;

namespace AccuX.Modules.BasicFinance.UI
{
    /// <summary>
    /// 小数位输入窗口（规格 §20 自定义舍入窗口）。
    /// WPF 只做输入与校验，不包含任何财务计算逻辑。
    /// </summary>
    public partial class RoundingWindow : Window
    {
        private readonly TextBox _digitsBox;

        public RoundingWindow(int defaultDigits, IntPtr ownerHandle)
        {
            Title = "一键舍入 - 小数位";
            Width = 300;
            Height = 170;
            ResizeMode = ResizeMode.NoResize;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ShowInTaskbar = false;
            WindowHelper.AttachOwner(this, ownerHandle);

            var panel = new StackPanel { Margin = new Thickness(16) };

            panel.Children.Add(new TextBlock
            {
                Text = "请输入保留的小数位数（0 - 15）：",
                Margin = new Thickness(0, 0, 0, 8)
            });

            _digitsBox = new TextBox
            {
                Text = defaultDigits.ToString(CultureInfo.InvariantCulture),
                Padding = new Thickness(4),
                Margin = new Thickness(0, 0, 0, 12)
            };
            panel.Children.Add(_digitsBox);

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var okButton = new Button
            {
                Content = "确定",
                Width = 72,
                Height = 26,
                IsDefault = true,
                Margin = new Thickness(0, 0, 8, 0)
            };
            okButton.Click += OnOk;

            var cancelButton = new Button
            {
                Content = "取消",
                Width = 72,
                Height = 26,
                IsCancel = true
            };

            buttons.Children.Add(okButton);
            buttons.Children.Add(cancelButton);
            panel.Children.Add(buttons);

            Content = panel;
            _digitsBox.Focus();
            _digitsBox.SelectAll();
        }

        public int Digits { get; private set; }

        private void OnOk(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(_digitsBox.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var digits)
                || digits < 0
                || digits > 15)
            {
                MessageBox.Show(this, "请输入 0 到 15 之间的整数。", "AccuX", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Digits = digits;
            DialogResult = true;
        }
    }
}
