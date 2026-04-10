using SqlServerTool.Helpers;
using SqlServerTool.Models;
using SqlServerTool.ViewModels;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using WpfBinding  = System.Windows.Data.Binding;
using WpfBrushes  = System.Windows.Media.Brushes;
using WpfTextBox  = System.Windows.Controls.TextBox;
using WpfCheckBox = System.Windows.Controls.CheckBox;

namespace SqlServerTool.Views
{
    public partial class EditRowDialog : Window
    {
        public EditRowViewModel ViewModel { get; }

        private readonly Action<EditRowViewModel>? _insertAction;

        /// <summary>
        /// 編集用コンストラクタ。
        /// insertAction を渡すと「新規追加」ボタンが有効になる。
        /// </summary>
        public EditRowDialog(EditRowViewModel viewModel, Action<EditRowViewModel>? insertAction = null)
        {
            InitializeComponent();
            ViewModel      = viewModel;
            _insertAction  = insertAction;
            DataContext    = ViewModel;

            BuildFieldsLayout();
        }

        // ─── フィールドレイアウト動的生成 ─────────────────────────────────────

        private static readonly InvertBoolConverter _invertBool = new();

        /// <summary>
        /// Fields リストから Grid を生成してコンテンツエリアに配置する。
        /// MAX 型フィールドがある場合はその行を * 高さ（伸長可能）にする。
        /// MAX フィールドがない場合は ScrollViewer で包む。
        /// </summary>
        private void BuildFieldsLayout()
        {
            var fields = ViewModel.Fields;
            bool hasMax = fields.Any(f => f.IsMaxType && !f.IsIdentity);

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(60) });

            for (int i = 0; i < fields.Count; i++)
            {
                var f = fields[i];
                bool isExpandable = f.IsMaxType && !f.IsIdentity;

                grid.RowDefinitions.Add(new RowDefinition
                {
                    Height    = isExpandable
                        ? new GridLength(1, GridUnitType.Star)
                        : GridLength.Auto,
                    MinHeight = isExpandable ? 60 : 0
                });

                // ── ラベル列 ──────────────────────────────────────────────────
                var label = new StackPanel
                {
                    VerticalAlignment = VerticalAlignment.Top,
                    Margin = new Thickness(0, 6, 0, 4)
                };
                label.Children.Add(new TextBlock
                    { Text = f.ColumnName, FontWeight = FontWeights.SemiBold });
                if (f.HasLogicalName)
                    label.Children.Add(new TextBlock
                        { Text = f.LogicalName, Foreground = WpfBrushes.DarkBlue, FontSize = 10 });
                label.Children.Add(new TextBlock
                    { Text = f.DataType, Foreground = WpfBrushes.Gray, FontSize = 10 });
                if (f.IsPrimaryKey)
                    label.Children.Add(new TextBlock
                        { Text = "PK", Foreground = WpfBrushes.DarkBlue,
                          FontSize = 10, FontWeight = FontWeights.Bold });
                if (f.IsIdentity)
                    label.Children.Add(new TextBlock
                        { Text = "IDENTITY", Foreground = WpfBrushes.Gray,
                          FontSize = 10, FontStyle = FontStyles.Italic });
                Grid.SetRow(label, i); Grid.SetColumn(label, 0);
                grid.Children.Add(label);

                // ── 入力欄 ────────────────────────────────────────────────────
                if (f.IsIdentity)
                {
                    var auto = new TextBlock
                    {
                        Text              = "（自動採番 - 入力不要）",
                        Foreground        = WpfBrushes.Gray,
                        FontStyle         = FontStyles.Italic,
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin            = new Thickness(4, 0, 0, 0)
                    };
                    Grid.SetRow(auto, i); Grid.SetColumn(auto, 1);
                    Grid.SetColumnSpan(auto, 2);
                    grid.Children.Add(auto);
                }
                else
                {
                    var tb = new WpfTextBox
                    {
                        AcceptsReturn             = true,
                        TextWrapping              = TextWrapping.Wrap,
                        VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                        Margin            = new Thickness(4, 4, 0, 4),
                        VerticalAlignment = isExpandable
                            ? VerticalAlignment.Stretch
                            : VerticalAlignment.Top,
                    };
                    if (!isExpandable)
                    {
                        tb.MinHeight = 28;
                        tb.MaxHeight = 80;
                    }
                    tb.SetBinding(WpfTextBox.TextProperty, new WpfBinding("Value")
                        { Source = f, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged });
                    tb.SetBinding(UIElement.IsEnabledProperty, new WpfBinding("IsNull")
                        { Source = f, Converter = _invertBool });

                    Grid.SetRow(tb, i); Grid.SetColumn(tb, 1);
                    grid.Children.Add(tb);

                    var cb = new WpfCheckBox
                    {
                        Content           = "NULL",
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin            = new Thickness(4, 0, 0, 0)
                    };
                    cb.SetBinding(WpfCheckBox.IsCheckedProperty,
                        new WpfBinding("IsNull") { Source = f });
                    Grid.SetRow(cb, i); Grid.SetColumn(cb, 2);
                    grid.Children.Add(cb);
                }
            }

            // MAX フィールドなし → ScrollViewer で包む（従来動作）
            if (hasMax)
            {
                FieldsGrid.Children.Add(grid);
            }
            else
            {
                var sv = new ScrollViewer
                    { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
                sv.Content = grid;
                FieldsGrid.Children.Add(sv);
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel.IsNewRecord && !ValidateFields()) return;
            DialogResult = true;
            Close();
        }

        private void NewRowButton_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateFields()) return;

            try
            {
                _insertAction!(ViewModel);
                // 成功したらダイアログを閉じず、フィールドをリセットして連続入力を可能にする
                foreach (var f in ViewModel.Fields.Where(f => !f.IsIdentity))
                {
                    f.IsNull = false;
                    f.Value  = string.Empty;
                }
                System.Windows.MessageBox.Show("新規行を追加しました。", "完了",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (SqlPreviewCancelledException)
            {
                // プレビューでキャンセル → ダイアログは開いたまま
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    $"追加エラー: {ex.Message}", "エラー",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        /// <summary>NOT NULL・非IDENTITY 列の未入力チェック。問題があれば確認を求め false を返す。</summary>
        private bool ValidateFields()
        {
            var emptyRequired = ViewModel.Fields
                .Where(f => !f.IsIdentity && !f.IsNull && string.IsNullOrEmpty(f.Value))
                .Select(f => f.ColumnName)
                .ToList();

            if (emptyRequired.Count == 0) return true;

            var result = System.Windows.MessageBox.Show(
                $"以下の列が空です。このまま続行しますか？\n\n{string.Join(", ", emptyRequired)}",
                "未入力確認",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            return result == MessageBoxResult.Yes;
        }
    }
}