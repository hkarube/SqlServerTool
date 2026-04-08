using SqlServerTool.Models;
using SqlServerTool.ViewModels;
using System;
using System.Linq;
using System.Windows;

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