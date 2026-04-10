using SqlServerTool.Models;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using System.Windows;

namespace SqlServerTool.Views
{
    public partial class CreateTableDialog : Window
    {
        private readonly ISet<string> _existingNames;
        private readonly ObservableCollection<EditableColumnInfo> _columns = new();

        public string NewTableName { get; private set; } = string.Empty;
        public IReadOnlyList<EditableColumnInfo> Columns => _columns;

        public CreateTableDialog(ISet<string> existingNames)
        {
            InitializeComponent();
            _existingNames    = existingNames;
            ColumnsGrid.ItemsSource = _columns;

            // 初期列を1つ追加
            _columns.Add(new EditableColumnInfo { ColumnName = "Id", DataType = "int", IsNullable = false, IsPrimaryKey = true });
        }

        private void AddColumnButton_Click(object sender, RoutedEventArgs e)
        {
            _columns.Add(new EditableColumnInfo { DataType = "nvarchar(50)", IsNullable = true });
            ColumnsGrid.ScrollIntoView(_columns[^1]);
        }

        private void RemoveColumnButton_Click(object sender, RoutedEventArgs e)
        {
            if (ColumnsGrid.SelectedItem is EditableColumnInfo col)
                _columns.Remove(col);
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
            => TryAccept();

        private void TryAccept()
        {
            // テーブル名バリデーション
            var tableName = TableNameBox.Text.Trim();
            var nameError = ValidateName(tableName);
            if (nameError != null) { ShowError(nameError); return; }

            // 列バリデーション
            if (_columns.Count == 0) { ShowError("列を1つ以上追加してください。"); return; }

            var emptyCol = _columns.FirstOrDefault(c => string.IsNullOrWhiteSpace(c.ColumnName));
            if (emptyCol != null) { ShowError("列名が空の行があります。"); return; }

            var emptyType = _columns.FirstOrDefault(c => string.IsNullOrWhiteSpace(c.DataType));
            if (emptyType != null) { ShowError($"「{emptyType.ColumnName}」のデータ型を指定してください。"); return; }

            var dupCol = _columns.GroupBy(c => c.ColumnName, StringComparer.OrdinalIgnoreCase)
                                 .FirstOrDefault(g => g.Count() > 1);
            if (dupCol != null) { ShowError($"列名「{dupCol.Key}」が重複しています。"); return; }

            ErrorText.Visibility = Visibility.Collapsed;
            NewTableName = tableName;
            DialogResult = true;
        }

        private string? ValidateName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "テーブル名を入力してください。";
            if (name.Length > 128)
                return "テーブル名は128文字以内にしてください。";
            if (!Regex.IsMatch(name, @"^[\p{L}_#@][\p{L}\p{N}_#@$]*$"))
                return "使用できない文字が含まれています。";
            if (_existingNames.Contains(name))
                return $"「{name}」はすでに存在します。別の名前を入力してください。";
            return null;
        }

        private void ShowError(string message)
        {
            ErrorText.Text       = message;
            ErrorText.Visibility = Visibility.Visible;
        }
    }
}
