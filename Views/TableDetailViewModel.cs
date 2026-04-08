using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SqlServerTool.Models;
using SqlServerTool.Services;
using SqlServerTool.ViewModels;
using System.Collections.ObjectModel;
using System.Data;

namespace SqlServerTool.ViewModels
{
    public partial class TableDetailViewModel : ObservableObject
    {
        private readonly SchemaService _schemaService;

        public string ObjectType { get; }   // TABLE / VIEW / PROCEDURE / FUNCTION

        [ObservableProperty] private string tableName     = string.Empty;
        [ObservableProperty] private string sqlText       = string.Empty;
        [ObservableProperty] private string rowCountText  = string.Empty;
        [ObservableProperty] private string statusMessage = string.Empty;

        // 構造タブ（読み取り用）
        public ObservableCollection<ColumnInfo> Columns { get; } = new();

        // 構造編集
        public ObservableCollection<EditableColumnInfo> EditableColumns { get; } = new();
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(IsNotEditingStructure))]
        private bool isEditingStructure = false;
        public bool IsNotEditingStructure => !IsEditingStructure;

        // コード編集（SP / Function）
        private string _originalSqlText = string.Empty;
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(IsNotEditingCode))]
        private bool isEditingCode = false;
        public bool IsNotEditingCode => !IsEditingCode;

        // データタブ
        private DataTable? _dataTable;
        public DataTable? DataRows
        {
            get => _dataTable;
            private set => SetProperty(ref _dataTable, value);
        }

        // 削除・更新用に元の型付きデータを保持
        private DataTable? _originalDataTable;

        // セッション内 SQL ログ（最新が先頭）
        public ObservableCollection<SessionLogEntry> SessionLog { get; } = new();

        // ── SQL プレビュー（コードビハインドから設定） ─────────────────────────
        /// <summary>
        /// SQL 実行前のプレビューダイアログ。(operationType, sql) → true=実行, false=キャンセル
        /// null の場合はダイアログなしで即実行。
        /// </summary>
        public Func<string, string, bool>? PreviewSql { get; set; }

        // 表示制御
        public bool IsTableOrView => ObjectType is "TABLE" or "VIEW";
        public bool IsCodeObject  => ObjectType is "PROCEDURE" or "FUNCTION";

        public TableDetailViewModel(string name, string objectType, SchemaService schemaService)
        {
            _schemaService = schemaService;
            TableName  = name;
            ObjectType = objectType;

            if (IsTableOrView)
            {
                LoadColumns();
                SqlText = $"SELECT TOP 500 * FROM [{name}]";
            }
            else
            {
                SqlText = _schemaService.GetSqlDefinition(name);
            }
        }

        private void LoadColumns()
        {
            Columns.Clear();
            try
            {
                foreach (var col in _schemaService.GetColumns(TableName))
                    Columns.Add(col);
            }
            catch (Exception ex)
            {
                StatusMessage = $"カラム取得エラー: {ex.Message}";
            }
        }

        [RelayCommand]
        private void SearchData()
        {
            if (!IsTableOrView) return;
            try
            {
                var (data, original, total) = _schemaService.GetTableData(TableName, SqlText);
                _originalDataTable = original;
                DataRows     = data;
                RowCountText = $"表示: {data.Rows.Count:#,##0} 件 / 全件: {total:#,##0} 件";
                StatusMessage = string.Empty;
            }
            catch (Exception ex)
            {
                RowCountText  = string.Empty;
                StatusMessage = $"エラー: {ex.Message}";
            }
        }

        [RelayCommand]
        private void CopyWithHeader()
        {
            if (DataRows == null || DataRows.Rows.Count == 0) return;
            var sb = new System.Text.StringBuilder();
            sb.AppendLine(string.Join("\t",
                DataRows.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
            foreach (DataRow row in DataRows.Rows)
                sb.AppendLine(string.Join("\t",
                    DataRows.Columns.Cast<DataColumn>().Select(c => row[c]?.ToString() ?? "")));
            System.Windows.Clipboard.SetText(sb.ToString());
            StatusMessage = "ヘッダ付きでクリップボードにコピーしました";
            System.Windows.MessageBox.Show("クリップボードにコピーしました", "完了",
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
        }

        // ─── コード編集（SP / Function） ─────────────────────────────────

        [RelayCommand]
        private void StartEditCode()
        {
            _originalSqlText = SqlText;
            IsEditingCode    = true;
            StatusMessage    = "SQLを編集中です。保存時に構文チェックを行います。";
        }

        [RelayCommand]
        private void CancelEditCode()
        {
            SqlText       = _originalSqlText;
            IsEditingCode = false;
            StatusMessage = string.Empty;
        }

        [RelayCommand]
        private async Task SaveCode()
        {
            // CREATE → ALTER に変換
            var alterSql = System.Text.RegularExpressions.Regex.Replace(
                SqlText,
                @"\bCREATE\s+(PROCEDURE|PROC|FUNCTION)\b",
                m => $"ALTER {m.Groups[1].Value}",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            StatusMessage = "構文チェック中...";

            var syntaxError = await Task.Run(() => _schemaService.CheckSqlSyntax(alterSql));
            if (syntaxError != null)
            {
                StatusMessage = $"構文エラー: {syntaxError}";
                return;
            }

            // SQL プレビュー
            if (PreviewSql != null && !PreviewSql("CODE", alterSql))
            {
                StatusMessage = "キャンセルしました";
                return;
            }

            try
            {
                await Task.Run(() => _schemaService.ExecuteAlterObject(alterSql));
                LogSql("CODE", alterSql);
                IsEditingCode = false;
                StatusMessage = "保存しました";
            }
            catch (Exception ex)
            {
                StatusMessage = $"保存エラー: {ex.Message}";
            }
        }

        // ─── 構造編集 ─────────────────────────────────────────────────────

        [RelayCommand]
        private void StartEditStructure()
        {
            EditableColumns.Clear();
            foreach (var c in Columns)
                EditableColumns.Add(new EditableColumnInfo(c));
            IsEditingStructure = true;
            StatusMessage = "構造を編集中です。変更後「保存」を押してください。";
        }

        /// <summary>選択行の上に新規列を挿入する（-1 なら末尾追加）</summary>
        public void InsertColumn(int insertIndex)
        {
            int n = 1;
            string name;
            do { name = $"NewColumn{n++}"; }
            while (EditableColumns.Any(c => c.ColumnName == name));

            var newCol = new EditableColumnInfo
            {
                ColumnName = name,
                DataType   = "nvarchar(50)",
                IsNullable = true
            };

            if (insertIndex >= 0 && insertIndex < EditableColumns.Count)
                EditableColumns.Insert(insertIndex, newCol);
            else
                EditableColumns.Add(newCol);
        }

        [RelayCommand]
        private void RemoveColumn(EditableColumnInfo? col)
        {
            if (col != null) EditableColumns.Remove(col);
        }

        [RelayCommand]
        private async Task SaveStructure()
        {
            var originalCols = Columns.ToList();
            var editedCols   = EditableColumns.ToList();

            try
            {
                // BuildAlterTableSql も try 内に入れて例外を確実にキャッチ
                var previewSqlText = _schemaService.BuildAlterTableSql(TableName, originalCols, editedCols);
                if (PreviewSql != null && !PreviewSql("STRUCTURE", previewSqlText))
                {
                    StatusMessage = "キャンセルしました";
                    return;
                }

                await Task.Run(() => _schemaService.AlterTableStructure(TableName, originalCols, editedCols));
                LogSql("STRUCTURE", previewSqlText);
                LoadColumns();
                IsEditingStructure = false;
                StatusMessage = "構造を更新しました";
            }
            catch (Exception ex)
            {
                StatusMessage = $"構造更新エラー: {ex.Message}";
                System.Windows.MessageBox.Show(
                    $"構造更新に失敗しました。\n\n{ex.Message}",
                    "エラー",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
            }
        }

        [RelayCommand]
        private void CancelEditStructure()
        {
            IsEditingStructure = false;
            StatusMessage = string.Empty;
        }

        // ─── 行削除 ───────────────────────────────────────────────────────

        /// <summary>指定行を DELETE する（コードビハインドから呼ぶ）</summary>
        public void DeleteRows(IEnumerable<System.Data.DataRowView> rows)
        {
            if (_originalDataTable == null)
                throw new InvalidOperationException("先にデータを検索してください。");

            var pkColumnNames = Columns.Where(c => c.IsPrimaryKey).Select(c => c.ColumnName).ToList();
            bool hasPk = pkColumnNames.Count > 0;

            var rowList = rows.ToList();

            // WHERE 条件を事前に計算（プレビュー用 + 実行用）
            var deleteParams = rowList.Select(row =>
            {
                int idx = DataRows!.Rows.IndexOf(row.Row);
                if (idx < 0) throw new InvalidOperationException("削除対象の行データが見つかりません。");
                var originalRow = _originalDataTable.Rows[idx];

                IEnumerable<(string Name, object? Value)> whereValues = hasPk
                    ? pkColumnNames.Select(name =>
                        (name, originalRow[name] == DBNull.Value ? (object?)null : originalRow[name]))
                    : _originalDataTable.Columns.Cast<DataColumn>()
                        .Select(c => (c.ColumnName,
                            originalRow[c] == DBNull.Value ? (object?)null : originalRow[c]));

                return whereValues.ToList();
            }).ToList();

            // SQL プレビュー（全行分を連結して表示）
            if (PreviewSql != null)
            {
                var previewSqlText = string.Join("\n\n", deleteParams.Select(wp =>
                    _schemaService.BuildDeleteSql(TableName, wp, !hasPk)));
                if (!PreviewSql("DELETE", previewSqlText))
                    throw new SqlPreviewCancelledException();
            }

            // 実行
            var sqlForLog = new System.Text.StringBuilder();
            for (int i = 0; i < rowList.Count; i++)
            {
                var wp  = deleteParams[i];
                var sql = _schemaService.BuildDeleteSql(TableName, wp, !hasPk);
                _schemaService.DeleteRow(TableName, wp, !hasPk);
                sqlForLog.AppendLine(sql);
            }
            LogSql("DELETE", sqlForLog.ToString().TrimEnd());
        }

        /// <summary>新規行ダイアログから呼ばれる INSERT 実行</summary>
        public void ExecuteInsert(EditRowViewModel editVm)
        {
            var insertCols = editVm.Fields
                .Where(f => !f.IsIdentity)
                .Select(f => (f.ColumnName, f.IsNull, f.Value))
                .ToList();

            var displaySql = _schemaService.BuildInsertSql(TableName, insertCols);

            if (PreviewSql != null && !PreviewSql("INSERT", displaySql))
                throw new SqlPreviewCancelledException();

            _schemaService.InsertRow(TableName, insertCols);
            LogSql("INSERT", displaySql);
        }

        /// <summary>行編集ダイアログから呼ばれる UPDATE 実行</summary>
        public void ExecuteUpdate(EditRowViewModel editVm, int originalRowIndex)
        {
            if (_originalDataTable == null)
                throw new InvalidOperationException("先にデータを検索してください。");

            var pkColumnNames = Columns.Where(c => c.IsPrimaryKey).Select(c => c.ColumnName).ToList();
            bool hasPk = pkColumnNames.Count > 0;

            var setColumns = editVm.Fields
                .Where(f => !f.IsPrimaryKey)
                .Select(f => (f.ColumnName, f.IsNull, f.Value))
                .ToList();

            var originalRow = _originalDataTable.Rows[originalRowIndex];
            List<(string Name, object? Value)> whereColumns;
            if (hasPk)
            {
                whereColumns = pkColumnNames.Select(name =>
                    (name, originalRow[name] == DBNull.Value ? (object?)null : originalRow[name])).ToList();
            }
            else
            {
                whereColumns = _originalDataTable.Columns.Cast<DataColumn>()
                    .Select(c => (c.ColumnName,
                        originalRow[c] == DBNull.Value ? (object?)null : originalRow[c])).ToList();
            }

            var displaySql = _schemaService.BuildUpdateSql(TableName, setColumns, whereColumns, !hasPk);

            if (PreviewSql != null && !PreviewSql("UPDATE", displaySql))
                throw new SqlPreviewCancelledException();

            _schemaService.UpdateRow(TableName, setColumns, whereColumns, !hasPk);
            LogSql("UPDATE", displaySql);
        }

        // ─── ログ ──────────────────────────────────────────────────────────

        private void LogSql(string opType, string sql)
        {
            SqlHistoryService.Instance.Add(new SqlHistoryEntry
            {
                ExecutedAt    = DateTime.Now,
                OperationType = opType,
                ObjectName    = TableName,
                Sql           = sql
            });
            System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                SessionLog.Insert(0, new SessionLogEntry(opType, sql)));
        }
    }

    public class SessionLogEntry
    {
        public string Time          { get; }
        public string OperationType { get; }
        public string Sql           { get; }
        public string Display       { get; }

        public SessionLogEntry(string opType, string sql)
        {
            Time          = DateTime.Now.ToString("HH:mm:ss");
            OperationType = opType;
            Sql           = sql;
            var firstLine = sql.Split('\n')
                .FirstOrDefault(l => !l.TrimStart().StartsWith("--") && l.Trim().Length > 0)
                ?? sql;
            Display = $"[{Time}] {opType}: {firstLine.Trim()}";
        }
    }
}
