using System;
using System.Windows;

namespace AccuX.Modules.BasicFinance.AmountConversion
{
    /// <summary>
    /// 金额折合参数窗口（规格 §20）。
    /// WPF 只做参数收集与校验，财务计算全部位于 <see cref="AmountConversionService"/>。
    /// </summary>
    public partial class AmountConversionView : Window
    {
        private readonly AmountConversionViewModel _viewModel;

        public AmountConversionView(IntPtr ownerHandle, AmountConversionViewModel viewModel = null)
        {
            InitializeComponent();

            _viewModel = viewModel ?? new AmountConversionViewModel();
            DataContext = _viewModel;

            UI.WindowHelper.AttachOwner(this, ownerHandle);
        }

        /// <summary>用户确认后的选项；取消时为 null。</summary>
        public AmountConversionOptions Options { get; private set; }

        private void OnOk(object sender, RoutedEventArgs e)
        {
            var options = _viewModel.BuildOptions();
            if (options == null)
            {
                return;
            }

            Options = options;
            DialogResult = true;
        }
    }
}
