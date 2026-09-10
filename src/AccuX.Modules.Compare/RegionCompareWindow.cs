using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using AccuX.Core.Cells;
using AccuX.Core.Logging;
using AccuX.Core.Operations;

namespace AccuX.Modules.Compare
{
    /// <summary>
    /// 单实例非模态区域对比窗口。界面只负责捕获目标和呈现结果，比较算法位于 RegionCompareService。
    /// </summary>
    public sealed class RegionCompareWindow : Window
    {
        private readonly IRegionCompareHost _host;
        private readonly RegionCompareService _service;
        private readonly CompareConfig _config;
        private readonly ILogger _logger;
        private readonly TextBlock _firstTargetText;
        private readonly TextBlock _secondTargetText;
        private readonly CheckBox _excludeHeaderBox;
        private readonly ListBox _firstOnlyList;
        private readonly ListBox _secondOnlyList;
        private readonly ListBox _sameList;
        private readonly TextBlock _summaryText;
        private RangeTarget _firstTarget;
        private RangeTarget _secondTarget;
        private RegionCompareReadState _state;
        private readonly List<RegionCompareCellPosition> _markedFirst = new List<RegionCompareCellPosition>();
        private readonly List<RegionCompareCellPosition> _markedSecond = new List<RegionCompareCellPosition>();
        private RangeTarget _markedFirstTarget;
        private RangeTarget _markedSecondTarget;

        public RegionCompareWindow(
            IRegionCompareHost host,
            RegionCompareService service,
            CompareConfig config,
            ILogger logger,
            IntPtr? ownerHandle = null)
        {
            _host = host ?? throw new ArgumentNullException(nameof(host));
            _service = service ?? throw new ArgumentNullException(nameof(service));
            _config = config ?? new CompareConfig();
            _logger = logger ?? NullLogger.Instance;

            Title = "区域对比";
            Width = 980;
            Height = 650;
            MinWidth = 760;
            MinHeight = 480;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ShowInTaskbar = false;
            if (ownerHandle.HasValue && ownerHandle.Value != IntPtr.Zero)
            {
                CompareWindowHelper.AttachOwner(this, ownerHandle.Value);
            }

            var root = new Grid { Margin = new Thickness(14) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(44) });

            var selectionPanel = new GroupBox { Header = "区域", Padding = new Thickness(10), Margin = new Thickness(0, 0, 0, 10) };
            var selectionGrid = new Grid();
            selectionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            selectionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            selectionGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            selectionGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            selectionGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            selectionGrid.Children.Add(new TextBlock { Text = "区域1：", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 8, 4) });
            _firstTargetText = CreateTargetText();
            Grid.SetColumn(_firstTargetText, 0); Grid.SetRow(_firstTargetText, 1); selectionGrid.Children.Add(_firstTargetText);
            var firstButton = new Button { Content = "使用当前选区", Width = 108, Height = 28, Margin = new Thickness(0, 6, 0, 0), HorizontalAlignment = HorizontalAlignment.Left };
            firstButton.Click += (s, e) => CaptureTarget(true);
            Grid.SetRow(firstButton, 2); selectionGrid.Children.Add(firstButton);

            var secondLabel = new TextBlock { Text = "区域2：", FontWeight = FontWeights.Bold, Margin = new Thickness(18, 0, 8, 4) };
            Grid.SetColumn(secondLabel, 1); selectionGrid.Children.Add(secondLabel);
            _secondTargetText = CreateTargetText(); Grid.SetColumn(_secondTargetText, 1); Grid.SetRow(_secondTargetText, 1); selectionGrid.Children.Add(_secondTargetText);
            var secondButton = new Button { Content = "使用当前选区", Width = 108, Height = 28, Margin = new Thickness(18, 6, 0, 0), HorizontalAlignment = HorizontalAlignment.Left };
            secondButton.Click += (s, e) => CaptureTarget(false);
            Grid.SetColumn(secondButton, 1); Grid.SetRow(secondButton, 2); selectionGrid.Children.Add(secondButton);
            selectionPanel.Content = selectionGrid;
            Grid.SetRow(selectionPanel, 0); root.Children.Add(selectionPanel);

            var options = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 10) };
            _excludeHeaderBox = new CheckBox { Content = "区域包含标题（各区域第一行不参与对比）", VerticalAlignment = VerticalAlignment.Center };
            _excludeHeaderBox.Checked += (s, e) => InvalidateResults();
            _excludeHeaderBox.Unchecked += (s, e) => InvalidateResults();
            options.Children.Add(_excludeHeaderBox);
            Grid.SetRow(options, 1); root.Children.Add(options);

            var resultGrid = new Grid();
            resultGrid.ColumnDefinitions.Add(new ColumnDefinition());
            resultGrid.ColumnDefinitions.Add(new ColumnDefinition());
            resultGrid.ColumnDefinitions.Add(new ColumnDefinition());
            _firstOnlyList = CreateResultList("区域1独有", 0, resultGrid);
            _secondOnlyList = CreateResultList("区域2独有", 1, resultGrid);
            _sameList = CreateResultList("相同项", 2, resultGrid);
            Grid.SetRow(resultGrid, 2); root.Children.Add(resultGrid);

            var footer = new Border
            {
                Margin = new Thickness(0, 8, 0, 0),
                Padding = new Thickness(0, 4, 0, 0),
                BorderThickness = new Thickness(0, 1, 0, 0),
                BorderBrush = System.Windows.Media.Brushes.LightGray,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            };
            var footerGrid = new Grid();
            footerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            footerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            _summaryText = new TextBlock { Text = "请分别捕获两个区域后点击“对比”。", VerticalAlignment = VerticalAlignment.Center, TextWrapping = TextWrapping.Wrap };
            Grid.SetColumn(_summaryText, 0); footerGrid.Children.Add(_summaryText);
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Center };
            buttons.Children.Add(CreateButton("对比", Compare));
            buttons.Children.Add(CreateButton("标记不同", MarkDifferent));
            buttons.Children.Add(CreateButton("标记相同", MarkSame));
            buttons.Children.Add(CreateButton("清除标记", ClearMarks));
            buttons.Children.Add(CreateButton("导出结果", Export));
            Grid.SetColumn(buttons, 1); footerGrid.Children.Add(buttons);
            footer.Child = footerGrid;
            Grid.SetRow(footer, 3); root.Children.Add(footer);
            Content = root;
        }

        private TextBlock CreateTargetText()
        {
            return new TextBlock { Text = "未选择", Foreground = System.Windows.Media.Brushes.DimGray, TextWrapping = TextWrapping.Wrap, MinHeight = 38 };
        }

        private ListBox CreateResultList(string title, int column, Grid parent)
        {
            var group = new GroupBox { Header = title, Margin = new Thickness(column == 0 ? 0 : 6, 0, column == 2 ? 0 : 6, 0) };
            var list = new ListBox { MinHeight = 220, BorderThickness = new Thickness(1), Padding = new Thickness(3) };
            group.Content = list;
            Grid.SetColumn(group, column); parent.Children.Add(group);
            return list;
        }

        private static Button CreateButton(string text, RoutedEventHandler handler)
        {
            var button = new Button { Content = text, Width = 108, Height = 28, Margin = new Thickness(6, 0, 0, 0) };
            button.Click += handler;
            return button;
        }

        private void CaptureTarget(bool first)
        {
            try
            {
                var target = _host.CaptureCurrentTarget();
                if (!_host.ValidateTarget(target, false, out var message))
                {
                    ShowError(message);
                    return;
                }

                if (first && _markedFirst.Count > 0 && !ReferenceEquals(_firstTarget, target))
                {
                    ShowError("区域1已有本窗口标记，请先清除标记后再更换区域。");
                    return;
                }
                if (!first && _markedSecond.Count > 0 && !ReferenceEquals(_secondTarget, target))
                {
                    ShowError("区域2已有本窗口标记，请先清除标记后再更换区域。");
                    return;
                }

                if (first) _firstTarget = target; else _secondTarget = target;
                InvalidateResults();
                var text = FormatTarget(target);
                if (first) _firstTargetText.Text = text; else _secondTargetText.Text = text;
                _summaryText.Text = "两个区域均捕获后点击“对比”。";
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
        }

        private void Compare(object sender, RoutedEventArgs e)
        {
            if (_firstTarget == null || _secondTarget == null)
            {
                ShowError("请先分别捕获区域1和区域2。");
                return;
            }

            try
            {
                if (!_host.ValidateTarget(_firstTarget, false, out var firstMessage)) { ShowError(firstMessage); return; }
                if (!_host.ValidateTarget(_secondTarget, false, out var secondMessage)) { ShowError(secondMessage); return; }
                var total = _firstTarget.CellCount + _secondTarget.CellCount;
                if (total > _config.MaxProcessCells)
                {
                    ShowError("两个区域合计包含 " + total + " 个单元格，超过上限 " + _config.MaxProcessCells + "。");
                    return;
                }
                if (total > _config.LargeSelectionWarning)
                {
                    var choice = MessageBox.Show(this, "两个区域合计包含 " + total + " 个单元格，继续比较可能需要较长时间，是否继续？", "AccuX", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                    if (choice != MessageBoxResult.Yes) return;
                }

                var firstRead = _host.ReadRegion(_firstTarget);
                var secondRead = _host.ReadRegion(_secondTarget);
                var result = _service.Compare(firstRead, secondRead, _excludeHeaderBox.IsChecked == true);
                _state = new RegionCompareReadState(firstRead, secondRead, result);
                Render(result);
                _summaryText.Text = FormatSummary(result);
                _logger.Info("区域对比完成：区域1独有=" + result.FirstOnly.Count
                    + ",区域2独有=" + result.SecondOnly.Count + ",相同项=" + result.Same.Count);
            }
            catch (Exception ex)
            {
                _state = null;
                ShowError("对比失败：" + ex.Message);
            }
        }

        private void Render(RegionCompareResult result)
        {
            FillList(_firstOnlyList, result.FirstOnly);
            FillList(_secondOnlyList, result.SecondOnly);
            FillList(_sameList, result.Same);
        }

        private static void FillList(ListBox list, IReadOnlyList<RegionCompareItem> items)
        {
            list.Items.Clear();
            foreach (var item in items)
            {
                list.Items.Add(FormatItem(item));
            }
        }

        private static string FormatItem(RegionCompareItem item)
        {
            var first = item.FirstCount == 0 ? "-" : item.FirstCount.ToString(CultureInfo.InvariantCulture);
            var second = item.SecondCount == 0 ? "-" : item.SecondCount.ToString(CultureInfo.InvariantCulture);
            return item.DisplayValue + "  [" + TypeName(item.CellType) + "]  区域1:" + first + " 区域2:" + second
                + "\n  位置1: " + FormatPositions(item.FirstPositions) + "\n  位置2: " + FormatPositions(item.SecondPositions);
        }

        private static string FormatPositions(IReadOnlyList<RegionCompareCellPosition> positions)
        {
            if (positions == null || positions.Count == 0) return "-";
            return string.Join(", ", positions.Select(p => "R" + (p.Row + 1) + "C" + (p.Column + 1)));
        }

        private void MarkDifferent(object sender, RoutedEventArgs e)
        {
            if (!EnsureResult()) return;
            ApplyMarks(_state.Result.FirstOnly.SelectMany(i => i.FirstPositions).ToList(), _state.Result.SecondOnly.SelectMany(i => i.SecondPositions).ToList(), true);
        }

        private void MarkSame(object sender, RoutedEventArgs e)
        {
            if (!EnsureResult()) return;
            ApplyMarks(_state.Result.Same.SelectMany(i => i.FirstPositions).ToList(), _state.Result.Same.SelectMany(i => i.SecondPositions).ToList(), false);
        }

        private void ApplyMarks(
            IReadOnlyList<RegionCompareCellPosition> firstPositions,
            IReadOnlyList<RegionCompareCellPosition> secondPositions,
            bool different)
        {
            if (firstPositions.Count == 0 && secondPositions.Count == 0)
            {
                ShowError("当前结果没有可标记的位置。");
                return;
            }

            if (firstPositions.Count > 0 && !_host.ValidateTarget(_firstTarget, true, out var firstMessage)) { ShowError(firstMessage); return; }
            if (secondPositions.Count > 0 && !_host.ValidateTarget(_secondTarget, true, out var secondMessage)) { ShowError(secondMessage); return; }
            if (firstPositions.Count > 0 && !_host.ValidatePositions(_firstTarget, firstPositions, out firstMessage)) { ShowError(firstMessage); return; }
            if (secondPositions.Count > 0 && !_host.ValidatePositions(_secondTarget, secondPositions, out secondMessage)) { ShowError(secondMessage); return; }

            try
            {
                if (firstPositions.Count > 0)
                {
                    _markedFirstTarget = _firstTarget;
                    AddUnique(_markedFirst, firstPositions);
                    _host.ApplyBackgroundColor(_firstTarget, firstPositions, different ? _config.FirstOnlyColor : _config.SameColor);
                }
                if (secondPositions.Count > 0)
                {
                    _markedSecondTarget = _secondTarget;
                    AddUnique(_markedSecond, secondPositions);
                    _host.ApplyBackgroundColor(_secondTarget, secondPositions, different ? _config.SecondOnlyColor : _config.SameColor);
                }
                _summaryText.Text = different ? "已标记两侧独有项。" : "已标记相同项。";
                _logger.Info("区域对比标记完成：类型=" + (different ? "不同" : "相同")
                    + ",区域1=" + firstPositions.Count + ",区域2=" + secondPositions.Count);
            }
            catch (Exception ex)
            {
                ShowError("标记失败：" + ex.Message + " 已成功写入的位置仍可通过“清除标记”处理。");
            }
        }

        private void ClearMarks(object sender, RoutedEventArgs e)
        {
            if (_markedFirst.Count == 0 && _markedSecond.Count == 0)
            {
                _summaryText.Text = "本窗口没有可清除的标记。";
                return;
            }

            var firstTarget = _markedFirstTarget ?? _firstTarget;
            var secondTarget = _markedSecondTarget ?? _secondTarget;
            if (_markedFirst.Count > 0 && !_host.ValidateTarget(firstTarget, true, out var firstMessage)) { ShowError(firstMessage); return; }
            if (_markedSecond.Count > 0 && !_host.ValidateTarget(secondTarget, true, out var secondMessage)) { ShowError(secondMessage); return; }
            if (_markedFirst.Count > 0 && !_host.ValidatePositions(firstTarget, _markedFirst, out firstMessage)) { ShowError(firstMessage); return; }
            if (_markedSecond.Count > 0 && !_host.ValidatePositions(secondTarget, _markedSecond, out secondMessage)) { ShowError(secondMessage); return; }
            try
            {
                if (_markedFirst.Count > 0) _host.ClearBackgroundColor(firstTarget, _markedFirst);
                if (_markedSecond.Count > 0) _host.ClearBackgroundColor(secondTarget, _markedSecond);
                _markedFirst.Clear(); _markedSecond.Clear();
                _markedFirstTarget = null; _markedSecondTarget = null;
                _summaryText.Text = "已清除本窗口实际标记的位置（不会恢复原底色）。";
            }
            catch (Exception ex)
            {
                ShowError("清除标记失败：" + ex.Message);
            }
        }

        private void Export(object sender, RoutedEventArgs e)
        {
            if (!EnsureResult()) return;
            try
            {
                var rows = new List<RegionCompareExportRow>();
                AddExportRows(rows, "区域1独有", _state.Result.FirstOnly);
                AddExportRows(rows, "区域2独有", _state.Result.SecondOnly);
                AddExportRows(rows, "相同项", _state.Result.Same);
                _host.ExportResults(new RegionCompareExportData(rows, FormatSummary(_state.Result)));
                _summaryText.Text = "结果已导出到新工作簿。";
                _logger.Info("区域对比结果已导出：行数=" + rows.Count);
            }
            catch (Exception ex)
            {
                ShowError("导出失败：" + ex.Message);
            }
        }

        private void AddExportRows(List<RegionCompareExportRow> rows, string category, IReadOnlyList<RegionCompareItem> items)
        {
            foreach (var item in items)
            {
                rows.Add(new RegionCompareExportRow
                {
                    Category = category,
                    Value = item.DisplayValue,
                    CellType = TypeName(item.CellType),
                    FirstCount = item.FirstCount,
                    SecondCount = item.SecondCount,
                    FirstLocations = FormatPositions(item.FirstPositions),
                    SecondLocations = FormatPositions(item.SecondPositions),
                    FirstWorkbook = item.FirstCount > 0 ? _firstTarget.WorkbookKey : string.Empty,
                    FirstWorksheet = item.FirstCount > 0 ? _firstTarget.WorksheetName : string.Empty,
                    FirstRange = item.FirstCount > 0 ? _firstTarget.Address : string.Empty,
                    SecondWorkbook = item.SecondCount > 0 ? _secondTarget.WorkbookKey : string.Empty,
                    SecondWorksheet = item.SecondCount > 0 ? _secondTarget.WorksheetName : string.Empty,
                    SecondRange = item.SecondCount > 0 ? _secondTarget.Address : string.Empty
                });
            }
        }

        private bool EnsureResult()
        {
            if (_state == null)
            {
                ShowError("请先点击“对比”生成结果。");
                return false;
            }
            if (!_host.ValidateTarget(_firstTarget, false, out var firstMessage)) { ShowError(firstMessage); return false; }
            if (!_host.ValidateTarget(_secondTarget, false, out var secondMessage)) { ShowError(secondMessage); return false; }
            try
            {
                var current = _service.Compare(
                    _host.ReadRegion(_firstTarget),
                    _host.ReadRegion(_secondTarget),
                    _excludeHeaderBox.IsChecked == true);
                if (!Equivalent(_state.Result, current))
                {
                    _state = null;
                    _firstOnlyList.Items.Clear(); _secondOnlyList.Items.Clear(); _sameList.Items.Clear();
                    ShowError("源区域内容或可见性已发生变化，请重新点击“对比”。");
                    return false;
                }
            }
            catch (Exception ex)
            {
                ShowError("重新验证区域失败：" + ex.Message);
                return false;
            }
            return true;
        }

        private static bool Equivalent(RegionCompareResult left, RegionCompareResult right)
        {
            return EquivalentItems(left.FirstOnly, right.FirstOnly)
                && EquivalentItems(left.SecondOnly, right.SecondOnly)
                && EquivalentItems(left.Same, right.Same);
        }

        private static bool EquivalentItems(IReadOnlyList<RegionCompareItem> left, IReadOnlyList<RegionCompareItem> right)
        {
            if (left.Count != right.Count) return false;
            for (var i = 0; i < left.Count; i++)
            {
                var a = left[i]; var b = right[i];
                if (!string.Equals(a.Key, b.Key, StringComparison.Ordinal)
                    || a.FirstCount != b.FirstCount || a.SecondCount != b.SecondCount
                    || !EquivalentPositions(a.FirstPositions, b.FirstPositions)
                    || !EquivalentPositions(a.SecondPositions, b.SecondPositions)) return false;
            }
            return true;
        }

        private static bool EquivalentPositions(IReadOnlyList<RegionCompareCellPosition> left, IReadOnlyList<RegionCompareCellPosition> right)
        {
            if (left.Count != right.Count) return false;
            for (var i = 0; i < left.Count; i++)
            {
                if (left[i].Row != right[i].Row || left[i].Column != right[i].Column) return false;
            }
            return true;
        }

        private void InvalidateResults()
        {
            _state = null;
            _firstOnlyList?.Items.Clear(); _secondOnlyList?.Items.Clear(); _sameList?.Items.Clear();
            if (_markedFirst.Count > 0 || _markedSecond.Count > 0)
            {
                _summaryText?.SetValue(TextBlock.TextProperty, "对比条件已改变，请先清除本窗口标记或重新对比。");
            }
        }

        private string FormatTarget(RangeTarget target)
        {
            return (target.WorkbookKey ?? string.Empty) + "\n" + target.WorksheetName + "!" + target.Address
                + "（" + target.CellCount.ToString(CultureInfo.InvariantCulture) + " 个单元格）";
        }

        private static string FormatSummary(RegionCompareResult result)
        {
            var s = result.Statistics;
            return string.Format(CultureInfo.CurrentCulture,
                "区域1独有 {0} 项，区域2独有 {1} 项，相同项 {2} 项；跳过隐藏 {3}/{4}、空白 {5}/{6}、不支持类型 {7}/{8}。",
                result.FirstOnly.Count, result.SecondOnly.Count, result.Same.Count,
                s.FirstHiddenCount, s.SecondHiddenCount, s.FirstBlankCount, s.SecondBlankCount,
                s.FirstUnsupportedCount, s.SecondUnsupportedCount);
        }

        private static string TypeName(CellValueType type)
        {
            switch (type)
            {
                case CellValueType.ConstantNumber:
                case CellValueType.FormulaNumber: return "数值";
                case CellValueType.Text:
                case CellValueType.FormulaText: return "文本";
                case CellValueType.Date:
                case CellValueType.FormulaDate: return "日期";
                case CellValueType.Boolean:
                case CellValueType.FormulaBoolean: return "逻辑值";
                case CellValueType.Error:
                case CellValueType.FormulaError: return "错误";
                default: return type.ToString();
            }
        }

        private static void AddUnique(List<RegionCompareCellPosition> target, IReadOnlyList<RegionCompareCellPosition> source)
        {
            foreach (var position in source)
            {
                if (!target.Any(existing => existing.Row == position.Row && existing.Column == position.Column)) target.Add(position);
            }
        }

        private void ShowError(string message)
        {
            MessageBox.Show(this, message ?? "操作失败。", "AccuX", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        private sealed class RegionCompareReadState
        {
            public RegionCompareReadState(RangeReadResult first, RangeReadResult second, RegionCompareResult result)
            {
                First = first; Second = second; Result = result;
            }

            public RangeReadResult First { get; }
            public RangeReadResult Second { get; }
            public RegionCompareResult Result { get; }
        }
    }

    internal static class CompareWindowHelper
    {
        public static void AttachOwner(Window window, IntPtr owner)
        {
            // Setting Owner through WPF is not possible for a native Excel HWND. The
            // window itself remains non-modal; this helper intentionally keeps the
            // host handle boundary in the UI assembly.
            window.SourceInitialized += (s, e) =>
            {
                try
                {
                    var helper = new System.Windows.Interop.WindowInteropHelper(window);
                    NativeMethods.SetWindowLongPtr(helper.Handle, -8, owner);
                }
                catch { }
            };
        }

        private static class NativeMethods
        {
            [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
            public static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);
        }
    }
}
