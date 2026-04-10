using SqlServerTool.Helpers;
using SqlServerTool.Models;
using SqlServerTool.ViewModels;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WpfBinding = System.Windows.Data.Binding;
using WpfMsg    = System.Windows.MessageBox;
using WpfMsgB   = System.Windows.MessageBoxButton;
using WpfMsgI   = System.Windows.MessageBoxImage;
using WpfMsgR   = System.Windows.MessageBoxResult;

namespace SqlServerTool.Views
{
    public partial class TableDetailWindow : Window
    {
        private readonly ObjectInfo          _sourceInfo;
        private readonly TableDetailViewModel _vm;

        public TableDetailWindow(TableDetailViewModel viewModel, ObjectInfo sourceInfo)
        {
            InitializeComponent();
            _sourceInfo = sourceInfo;
            _vm         = viewModel;
            DataContext = _vm;

            // AvalonEdit は DependencyProperty バインディング不可なので直接セット
            if (_vm.IsCodeObject)
                CodeEditor.Text = _vm.SqlText;
            else
                DataSqlEditor.Text = _vm.SqlText;

            // データSQLエディタの変更をViewModelに反映
            DataSqlEditor.TextChanged += (s, e) => _vm.SqlText = DataSqlEditor.Text;

            // コードエディタの変更をViewModelに反映
            CodeEditor.TextChanged += (s, e) => { if (_vm.IsEditingCode) _vm.SqlText = CodeEditor.Text; };

            // IsEditingCode / SqlText 変化でエディタの読み書き状態と内容を同期
            _vm.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(_vm.IsEditingCode))
                    CodeEditor.IsReadOnly = !_vm.IsEditingCode;
                if (e.PropertyName == nameof(_vm.SqlText) && _vm.IsCodeObject)
                    CodeEditor.Text = _vm.SqlText;
            };

            // Ctrl+Enter で検索実行
            DataSqlEditor.PreviewKeyDown += (s, e) =>
            {
                if (e.Key == Key.Enter && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
                {
                    _vm.SearchDataCommand.Execute(null);
                    e.Handled = true;
                }
            };

            // 結果グリッドのダブルクリックで行編集
            ResultGrid.MouseDoubleClick += (s, e) =>
            {
                if (ResultGrid.SelectedItem is DataRowView)
                    EditRowButton_Click(s, e);
            };

            // SQL プレビューダイアログを ViewModel に接続
            _vm.PreviewSql = (opType, sql) =>
            {
                var dlg = new SqlPreviewDialog(opType, _vm.TableName, sql) { Owner = this };
                return dlg.ShowDialog() == true;
            };

            _sourceInfo.IsOpen = true;

            // 列追加時に編集グリッドを挿入行へスクロール
            _vm.EditableColumns.CollectionChanged += (s, e) =>
            {
                if (e.NewItems?.Count > 0)
                    EditColumnsGrid.ScrollIntoView(e.NewItems[0]);
            };

            // SP/Function の場合は SQL定義タブ（index=2）を選択状態にする
            if (_vm.IsCodeObject)
                MainTab.SelectedIndex = 2;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _sourceInfo.IsOpen = false;
        }

        // ─── 結果グリッド：自動列生成 ─────────────────────────────────────────

        private static readonly NullBackgroundConverter _nullBg      = new();
        private static readonly DbNullDisplayConverter  _nullDisplay = new();

        /// <summary>
        /// "(NULL)" 表示・テキスト省略（1行）の ElementStyle を設定する。
        /// NULL 背景色は ResultGrid の CellStyle DataTrigger（XAML）で処理。
        /// </summary>
        private void ResultGrid_AutoGeneratingColumn(
            object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (e.Column is not DataGridTextColumn textCol) return;

            var colName = e.PropertyName;

            // ── TextBlock 書式（1行・省略表示） ───────────────────────────────
            var elemStyle = new Style(typeof(TextBlock));
            elemStyle.Setters.Add(new Setter(TextBlock.TextTrimmingProperty,
                TextTrimming.CharacterEllipsis));
            elemStyle.Setters.Add(new Setter(TextBlock.VerticalAlignmentProperty,
                VerticalAlignment.Center));
            textCol.ElementStyle = elemStyle;

            // ── "(NULL)" 文字列表示 ────────────────────────────────────────────
            textCol.Binding = new WpfBinding($"[{colName}]")
            {
                Converter = _nullDisplay,
                Mode      = System.Windows.Data.BindingMode.OneWay
            };
        }

        private void DeleteRowButton_Click(object sender, RoutedEventArgs e)
        {
            var selected = ResultGrid.SelectedItems.Cast<DataRowView>().ToList();
            if (selected.Count == 0)
            {
                WpfMsg.Show("削除する行を選択してください", "情報",
                    WpfMsgB.OK, WpfMsgI.Information);
                return;
            }

            var confirm = WpfMsg.Show(
                $"{selected.Count} 件を削除しますか？\nこの操作は元に戻せません。",
                "削除確認", WpfMsgB.YesNo, WpfMsgI.Warning);
            if (confirm != WpfMsgR.Yes) return;

            try
            {
                _vm.DeleteRows(selected);
                _vm.StatusMessage = $"{selected.Count} 件削除しました";
                _vm.SearchDataCommand.Execute(null);
            }
            catch (SqlPreviewCancelledException)
            {
                // ユーザーがプレビューでキャンセル → 何もしない
            }
            catch (Exception ex)
            {
                WpfMsg.Show($"削除エラー: {ex.Message}", "エラー",
                    WpfMsgB.OK, WpfMsgI.Error);
            }
        }

        private void AddColumnButton_Click(object sender, RoutedEventArgs e)
        {
            int insertIndex = EditColumnsGrid.SelectedIndex;
            _vm.InsertColumn(insertIndex);
        }

        private void RemoveColumnButton_Click(object sender, RoutedEventArgs e)
        {
            if (EditColumnsGrid.SelectedItem is EditableColumnInfo col)
                _vm.RemoveColumnCommand.Execute(col);
            else
                WpfMsg.Show("削除する列を選択してください。", "情報",
                    WpfMsgB.OK, WpfMsgI.Information);
        }

        private void NewRowButton_Click(object sender, RoutedEventArgs e)
        {
            if (_vm.Columns.Count == 0)
            {
                WpfMsg.Show("先に「▶ 実行」でデータを取得してください。", "情報",
                    WpfMsgB.OK, WpfMsgI.Information);
                return;
            }
            OpenNewRowDialog();
        }

        /// <summary>新規行ダイアログを開く</summary>
        internal void OpenNewRowDialog()
        {
            if (_vm.Columns.Count == 0) return;

            var rowData = _vm.Columns.ToDictionary(
                c => c.ColumnName,
                c => c.IsIdentity ? "(NULL)" : (c.IsNullable ? "(NULL)" : string.Empty));

            var editVm = new EditRowViewModel(_vm.TableName, _vm.Columns, rowData)
            {
                IsNewRecord = true
            };
            var dialog = new EditRowDialog(editVm) { Owner = this };
            dialog.Title = "新規行の作成";
            if (dialog.ShowDialog() != true) return;

            try
            {
                _vm.ExecuteInsert(editVm);
                _vm.StatusMessage = "新規行を追加しました";
                _vm.SearchDataCommand.Execute(null);
            }
            catch (SqlPreviewCancelledException) { }
            catch (Exception ex)
            {
                WpfMsg.Show($"追加エラー: {ex.Message}", "エラー",
                    WpfMsgB.OK, WpfMsgI.Error);
            }
        }

        // ─── SQL ログパネル ────────────────────────────────────────────────

        private void SqlLogList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
            => ShowLogEntry();

        private void SqlLogView_Click(object sender, RoutedEventArgs e)
            => ShowLogEntry();

        private void SqlLogCopy_Click(object sender, RoutedEventArgs e)
        {
            if (SqlLogList.SelectedItem is not SessionLogEntry entry) return;
            for (int i = 0; i < 3; i++)
            {
                try { System.Windows.Clipboard.SetDataObject(entry.Sql, true); return; }
                catch { System.Threading.Thread.Sleep(100); }
            }
        }

        private void OpenHistoryButton_Click(object sender, RoutedEventArgs e)
        {
            var win = new SqlHistoryWindow(null) { Owner = this };
            win.Show();
        }

        private void ShowLogEntry()
        {
            if (SqlLogList.SelectedItem is not SessionLogEntry entry) return;
            new SqlPreviewDialog(entry.OperationType, _vm.TableName, entry.Sql, viewOnly: true)
            {
                Owner = this
            }.ShowDialog();
        }

        private void EditRowButton_Click(object sender, RoutedEventArgs e)
        {
            if (ResultGrid.SelectedItem is not DataRowView rowView)
            {
                WpfMsg.Show("編集する行を選択してください", "情報",
                    WpfMsgB.OK, WpfMsgI.Information);
                return;
            }

            int rowIndex = _vm.DataRows!.Rows.IndexOf(rowView.Row);
            var rowData  = _vm.DataRows!.Columns.Cast<DataColumn>()
                .ToDictionary(
                    c => c.ColumnName,
                    c => rowView.Row[c]?.ToString() ?? "(NULL)");

            var editVm = new EditRowViewModel(_vm.TableName, _vm.Columns, rowData);
            var dialog = new EditRowDialog(editVm, vm =>
            {
                _vm.ExecuteInsert(vm);
                _vm.StatusMessage = "新規行を追加しました";
                _vm.SearchDataCommand.Execute(null);
            }) { Owner = this };
            dialog.ShowDialog();

            if (dialog.DialogResult != true) return;

            try
            {
                _vm.ExecuteUpdate(editVm, rowIndex);
                _vm.StatusMessage = "更新しました";
                _vm.SearchDataCommand.Execute(null);
            }
            catch (SqlPreviewCancelledException) { }
            catch (Exception ex)
            {
                WpfMsg.Show($"更新エラー: {ex.Message}", "エラー",
                    WpfMsgB.OK, WpfMsgI.Error);
            }
        }
    }
}
