using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace AccuX.Modules.BasicFinance.AmountConversion
{
    /// <summary>
    /// 金额折合窗口的 ViewModel（规格 §5.4 / §20）。
    /// 只负责选项绑定与基础校验，不包含财务计算逻辑。
    /// </summary>
    public sealed class AmountConversionViewModel : INotifyPropertyChanged
    {
        private ConversionMode _mode = ConversionMode.Divide;
        private decimal _rate = 10000m;
        private int _digits = 2;
        private bool _appendWanSuffix;
        private string _validationMessage;

        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>可选折合率预设：除百 / 除千 / 除万。</summary>
        public IReadOnlyList<RatePreset> Presets { get; } = new List<RatePreset>
        {
            new RatePreset("除百", 100m),
            new RatePreset("除千", 1000m),
            new RatePreset("除万", 10000m)
        };

        public ConversionMode Mode
        {
            get { return _mode; }
            set { SetField(ref _mode, value); }
        }

        public bool IsDivide
        {
            get { return _mode == ConversionMode.Divide; }
            set { Mode = value ? ConversionMode.Divide : ConversionMode.Multiply; }
        }

        public bool IsMultiply
        {
            get { return _mode == ConversionMode.Multiply; }
            set { Mode = value ? ConversionMode.Multiply : ConversionMode.Divide; }
        }

        public decimal Rate
        {
            get { return _rate; }
            set { SetField(ref _rate, value); }
        }

        public int Digits
        {
            get { return _digits; }
            set { SetField(ref _digits, value); }
        }

        public bool AppendWanSuffix
        {
            get { return _appendWanSuffix; }
            set { SetField(ref _appendWanSuffix, value); }
        }

        public string ValidationMessage
        {
            get { return _validationMessage; }
            private set { SetField(ref _validationMessage, value); }
        }

        /// <summary>
        /// 校验并生成业务选项；失败时返回 null 并设置 ValidationMessage。
        /// </summary>
        public AmountConversionOptions BuildOptions()
        {
            var options = new AmountConversionOptions
            {
                Mode = _mode,
                Rate = _rate,
                Digits = _digits,
                AppendWanSuffix = _appendWanSuffix
            };

            if (!AmountConversionOptions.TryValidate(options, out var error))
            {
                ValidationMessage = error;
                return null;
            }

            ValidationMessage = null;
            return options;
        }

        private void SetField<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
            {
                return;
            }

            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// 折合率预设项。
    /// </summary>
    public sealed class RatePreset
    {
        public RatePreset(string name, decimal rate)
        {
            Name = name;
            Rate = rate;
        }

        public string Name { get; }

        public decimal Rate { get; }

        public override string ToString()
        {
            return Name + "（" + Rate.ToString(CultureInfo.InvariantCulture) + "）";
        }
    }
}
