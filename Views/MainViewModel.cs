using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SqlServerTool.Models;
using SqlServerTool.Services;
using SqlServerTool.Views;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Windows.Forms;
using WpfApp  = System.Windows.Application;
using WpfClip = System.Windows.Clipboard;
using WpfMsg  = System.Windows.MessageBox;
using WpfMsgB = System.Windows.MessageBoxButton;
using WpfMsgI = System.Windows.MessageBoxImage;
using WpfMsgR = System.Windows.MessageBoxResult;
using WpfWin  = System.Windows.Window;
using WinSave = Microsoft.Win32.SaveFileDialog;
using WinOpen = Microsoft.Win32.OpenFileDialog;

namespace SqlServerTool.ViewModels
{
    // ─────────────────────────────────────────────────────────────────────────
    // タブ基底クラス
    // ─────────────────────────────────────────────────────────────────────────

    public abstract class TabBase : ObservableObject
    {
        public abstract string Header { get; }
        public virtual bool IsClosable  => false;
        public virtual bool IsAddButton => false;
    }

    /// <summary>オブジェクト一覧タブ（常時先頭、閉じ不可）</summary>
    public partial class ObjectListTab : TabBase
    {
        public override string Header => "一覧";
        public ObservableCollection<ObjectInfo> Objects { get; } = new();
    }

    /// <summary>SQL クエリタブ（閉じ可能）</summary>
    public partial class QueryTab : TabBase
    {
        public override string Header     => $"クエリ {Index}";
        public override bool   IsClosable => true;

        public int Index { get; }
        private readonly DatabaseService _dbService;
        private readonly SchemaService?  _schemaService;
        private readonly AppSettings?    _appSettings;

        [ObservableProperty] private string sqlText      = string.Empty;
        [ObservableProperty] private string statusMessage = string.Empty;

        /// <summary>全クエリ共通の実行ログ（SqlHistoryService のグローバル履歴）</summary>
        public ObservableCollection<SqlHistoryEntry> SqlLog
            => SqlHistoryService.Instance.Entries;

        private DataTable? _resultTable;
        public DataTable? ResultTable
        {
            get => _resultTable;
            private set => SetProperty(ref _resultTable, value);
        }

        public ObservableCollection<QueryTreeNode> TreeNodes { get; } = new();

        public QueryTab(int index, DatabaseService dbService,
                        SchemaService? schemaService = null, AppSettings? appSettings = null)
        {
            Index          = index;
            _dbService     = dbService;
            _schemaService = schemaService;
            _appSettings   = appSettings;
            BuildTree();
        }

        private void BuildTree()
        {
            if (_schemaService == null) return;
            TreeNodes.Clear();

            var categories = new[]
            {
                ("テーブル",             "Tables",      "TABLE"),
                ("ビュー",               "Views",       "VIEW"),
                ("ストアドプロシージャ", "StoredProcs", "PROCEDURE"),
                ("ファンクション",       "Functions",   "FUNCTION"),
            };

            foreach (var (label, cat, type) in categories)
            {
                var localCat  = cat;
                var localType = type;
                TreeNodes.Add(new QueryTreeNode(label, "category",
                    () => LoadObjectChildren(localCat, localType)));
            }
        }

        private IEnumerable<QueryTreeNode> LoadObjectChildren(string category, string objectType)
        {
            var names = category switch
            {
                "Tables"      => _schemaService!.GetTableNames(),
                "Views"       => _schemaService!.GetViewNames(),
                "StoredProcs" => _schemaService!.GetStoredProcNames(),
                "Functions"   => _schemaService!.GetFunctionNames(),
                _             => new List<string>()
            };

            foreach (var name in names)
            {
                var localName = name;
                if (objectType is "TABLE" or "VIEW")
                    yield return new QueryTreeNode(name, objectType,
                        () => LoadColumnChildren(localName));
                else
                    yield return new QueryTreeNode(name, objectType);
            }
        }

        private IEnumerable<QueryTreeNode> LoadColumnChildren(string tableName)
        {
            return _schemaService!.GetColumns(tableName)
                .Select(c => new QueryTreeNode(c.ColumnName, "column"));
        }

        [RelayCommand]
        public void ExecuteSql()
        {
            if (string.IsNullOrWhiteSpace(SqlText)) return;
            try
            {
                using var conn = _dbService.GetConnection();
                conn.Open();
                using var cmd = new Microsoft.Data.SqlClient.SqlCommand(SqlText, conn);
                cmd.CommandTimeout = 120;

                if (SqlText.TrimStart().StartsWith("SELECT", StringComparison.OrdinalIgnoreCase)
                 || SqlText.TrimStart().StartsWith("WITH",   StringComparison.OrdinalIgnoreCase))
                {
                    var adapter = new Microsoft.Data.SqlClient.SqlDataAdapter(cmd);
                    var dt = new DataTable();
                    adapter.Fill(dt);
                    // null → DataTable の場を設定し、DataGrid の AutoGenerateColumns が動くよう再設定
                    ResultTable   = null;
                    ResultTable   = dt;
                    StatusMessage = $"{dt.Rows.Count} 件取得しました";
                    SqlHistoryService.Instance.Add(new SqlHistoryEntry
                    {
                        ExecutedAt    = DateTime.Now,
                        OperationType = "SELECT",
                        ObjectName    = "クエリ",
                        Sql           = SqlText
                    });
                }
                else
                {
                    ResultTable   = null;
                    int affected  = cmd.ExecuteNonQuery();
                    StatusMessage = $"{affected} 件処理しました";
                    SqlHistoryService.Instance.Add(new SqlHistoryEntry
                    {
                        ExecutedAt    = DateTime.Now,
                        OperationType = "EXEC",
                        ObjectName    = "クエリ",
                        Sql           = SqlText
                    });
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"エラー: {ex.Message}";
            }
        }

        [RelayCommand]
        public void CopyWithHeader()
        {
            if (ResultTable == null || ResultTable.Rows.Count == 0) return;
            var sb = new System.Text.StringBuilder();
            sb.AppendLine(string.Join("\t",
                ResultTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
            foreach (DataRow row in ResultTable.Rows)
                sb.AppendLine(string.Join("\t",
                    ResultTable.Columns.Cast<DataColumn>().Select(c => row[c]?.ToString() ?? "")));
            WpfClip.SetText(sb.ToString());
            StatusMessage = "ヘッダ付きでクリップボードにコピーしました";
            WpfMsg.Show("クリップボードにコピーしました", "完了", WpfMsgB.OK, WpfMsgI.Information);
        }

        [RelayCommand]
        private void SaveSql()
        {
            if (string.IsNullOrWhiteSpace(SqlText)) return;

            // 保存時に最新設定を読み直す（タブ生成後に設定変更しても反映される）
            var latestSettings = SettingsService.Load();
            var workFolder = latestSettings.WorkFolder ?? string.Empty;
            var initDir = !string.IsNullOrEmpty(workFolder)
                ? workFolder
                : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            var dialog = new WinSave
            {
                Filter           = "SQL ファイル (*.sql)|*.sql|すべてのファイル (*.*)|*.*",
                DefaultExt       = "sql",
                InitialDirectory = initDir,
                FileName         = $"query_{DateTime.Now:yyyyMMdd_HHmm}.sql"
            };
            if (dialog.ShowDialog() != true) return;

            File.WriteAllText(dialog.FileName, SqlText, System.Text.Encoding.UTF8);
            StatusMessage = $"保存しました: {System.IO.Path.GetFileName(dialog.FileName)}";
        }
    }

    /// <summary>"+" タブ（新規クエリタブ追加ボタン）</summary>
    public class AddTabSentinel : TabBase
    {
        public override string Header     => "+";
        public override bool   IsAddButton => true;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // MainViewModel
    // ─────────────────────────────────────────────────────────────────────────

    public partial class MainViewModel : ObservableObject
    {
        private readonly DatabaseService _dbService;
        private SchemaService? _schemaService;
        private ScriptService? _scriptService;
        private BackupService? _backupService;
        private AppSettings    _appSettings;

        // ── タブ管琁E──────────────────────────────────────────────────────────
        public ObservableCollection<TabBase> Tabs { get; } = new();
        private readonly ObjectListTab  _listTab     = new();
        private readonly AddTabSentinel _addSentinel = new();
        private int _queryTabCount = 0;
        private bool _addingTab = false;

        private TabBase? _selectedTab;
        public TabBase? SelectedTab
        {
            get => _selectedTab;
            set
            {
                if (value is AddTabSentinel)
                {
                    // 二重起動防止ガード
                    if (!_addingTab)
                        WpfApp.Current.Dispatcher.InvokeAsync(AddNewQueryTab);
                    return;
                }
                SetProperty(ref _selectedTab, value);
            }
        }

        // ── 状慁E─────────────────────────────────────────────────────────────
        [ObservableProperty] private string connectionStatus = "未接続";
        [ObservableProperty] private bool   isConnected      = false;
        [ObservableProperty] private string currentCategory  = "Tables";

        public MainViewModel(DatabaseService dbService)
        {
            _dbService   = dbService;
            _appSettings = SettingsService.Load();

            Tabs.Add(_listTab);
            Tabs.Add(_addSentinel);
            _selectedTab = _listTab;

            // App.xaml.cs でログイン済みの場合はそのまま初期化する
            if (_dbService.IsConnected)
            {
                IsConnected      = true;
                ConnectionStatus = $"接続中: {_dbService.ServerName} / {_dbService.DatabaseName}";
                _schemaService   = new SchemaService(_dbService);
                _scriptService   = new ScriptService(_dbService, _schemaService);
                _backupService   = new BackupService(_dbService);
                SelectCategory("Tables");
            }
        }

        // ── 接綁E─────────────────────────────────────────────────────────────

        [RelayCommand]
        private void OpenConnection()
        {
            var vm     = new ConnectionViewModel(_dbService);
            var dialog = new ConnectionDialog(vm) { Owner = WpfApp.Current.MainWindow };
            dialog.ShowDialog();

            if (vm.IsConnected)
            {
                IsConnected      = true;
                ConnectionStatus = $"接続中: {_dbService.ServerName} / {_dbService.DatabaseName}";
                _schemaService   = new SchemaService(_dbService);
                _scriptService   = new ScriptService(_dbService, _schemaService);
                _backupService   = new BackupService(_dbService);
                SelectCategory("Tables");
            }
        }

        // ── 一覧の選択アイテム（メニューコマンド用） ──────────────────────────
        public List<ObjectInfo> SelectedObjectInfos { get; set; } = new();

        // ── 左ペインのカテゴリ選択 ────────────────────────────────────────────

        [RelayCommand]
        private void SelectCategory(string category)
        {
            if (!IsConnected) return;
            CurrentCategory = category;
            LoadObjectList(category);
            SelectedTab = _listTab;
        }

        // ── 設定ウィンドウ ────────────────────────────────────────────────────

        [RelayCommand]
        private void OpenSettings()
        {
            var vm  = new SettingsViewModel(_appSettings);
            var win = new SettingsWindow(vm) { Owner = WpfApp.Current.MainWindow };
            win.ShowDialog();
            // 設定が保存されたら再読み込み
            _appSettings = SettingsService.Load();
        }

        private void LoadObjectList(string category)
        {
            if (_schemaService == null) return;
            _listTab.Objects.Clear();
            try
            {
                foreach (var obj in _schemaService.GetObjectList(category))
                    _listTab.Objects.Add(obj);
            }
            catch (Exception ex)
            {
                WpfMsg.Show($"一覧取得エラー: {ex.Message}", "エラー", WpfMsgB.OK, WpfMsgI.Error);
            }
        }

        // ── 一覧更新 ─────────────────────────────────────────────────────────

        [RelayCommand]
        private void RefreshList()
        {
            if (IsConnected) LoadObjectList(CurrentCategory);
        }

        // ── オブジェクト削除 ─────────────────────────────────────────────────

        [RelayCommand]
        private void DeleteObject(System.Collections.IList? selectedItems)
        {
            if (_schemaService == null || selectedItems == null || selectedItems.Count == 0)
            {
                WpfMsg.Show("削除するオブジェクトを選択してください", "情報", WpfMsgB.OK, WpfMsgI.Information);
                return;
            }

            var targets = selectedItems.Cast<ObjectInfo>().ToList();
            var names   = string.Join("\n", targets.Select(t => $"  {t.ObjectType}: {t.Name}"));
            var confirm = WpfMsg.Show(
                $"以下のオブジェクトを削除しますか？\nこの操作は元に戻せません。\n\n{names}",
                "削除確認", WpfMsgB.YesNo, WpfMsgI.Warning);
            if (confirm != WpfMsgR.Yes) return;

            var errors = new List<string>();
            foreach (var t in targets)
            {
                try { _schemaService.DropObject(t.Name, t.ObjectType); }
                catch (Exception ex) { errors.Add($"{t.Name}: {ex.Message}"); }
            }

            if (errors.Count > 0)
                WpfMsg.Show($"削除エラー:\n{string.Join("\n", errors)}", "エラー", WpfMsgB.OK, WpfMsgI.Error);

            LoadObjectList(CurrentCategory);
        }

        // ── オブジェクト名変更 ────────────────────────────────────────────────

        [RelayCommand]
        private void RenameObject(System.Collections.IList? selectedItems)
        {
            if (_schemaService == null) return;

            if (selectedItems == null || selectedItems.Count == 0)
            {
                WpfMsg.Show("名前を変更するオブジェクトを選択してください。",
                    "情報", WpfMsgB.OK, WpfMsgI.Information);
                return;
            }
            if (selectedItems.Count > 1)
            {
                WpfMsg.Show("名前の変更は1件ずつ行ってください。",
                    "情報", WpfMsgB.OK, WpfMsgI.Information);
                return;
            }

            var item = (ObjectInfo)selectedItems[0]!;

            if (item.IsOpen)
            {
                WpfMsg.Show(
                    $"「{item.Name}」は現在詳細ウィンドウで開かれているため名前を変更できません。\n" +
                    "詳細ウィンドウを閉じてから操作してください。",
                    "編集", WpfMsgB.OK, WpfMsgI.Warning);
                return;
            }

            var existingNames = _listTab.Objects
                .Select(o => o.Name)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var dialog = new Views.RenameDialog(item.Name, existingNames)
            {
                Owner = WpfApp.Current.MainWindow
            };
            if (dialog.ShowDialog() != true) return;

            try
            {
                _schemaService.RenameObject(item.Name, dialog.NewName);
                LoadObjectList(CurrentCategory);
            }
            catch (Exception ex)
            {
                WpfMsg.Show($"名前変更エラー: {ex.Message}", "エラー", WpfMsgB.OK, WpfMsgI.Error);
            }
        }

        // ── テーブル新規作成 ──────────────────────────────────────────────────

        [RelayCommand]
        private void CreateTable()
        {
            if (_schemaService == null)
            {
                WpfMsg.Show("先に接続してください。", "情報", WpfMsgB.OK, WpfMsgI.Information);
                return;
            }

            if (CurrentCategory != "Tables")
            {
                WpfMsg.Show("テーブル一覧を表示中のときのみ新規作成できます。",
                    "情報", WpfMsgB.OK, WpfMsgI.Information);
                return;
            }

            var existingNames = _listTab.Objects
                .Select(o => o.Name)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var dialog = new Views.CreateTableDialog(existingNames)
            {
                Owner = WpfApp.Current.MainWindow
            };
            if (dialog.ShowDialog() != true) return;

            try
            {
                _schemaService.CreateTable(dialog.NewTableName, dialog.Columns);
                LoadObjectList(CurrentCategory);
            }
            catch (Exception ex)
            {
                WpfMsg.Show($"テーブル作成エラー: {ex.Message}", "エラー", WpfMsgB.OK, WpfMsgI.Error);
            }
        }

        // ── テーブルコピー ────────────────────────────────────────────────────

        [RelayCommand]
        private void CopyObject(System.Collections.IList? selectedItems)
        {
            if (_schemaService == null) return;

            if (selectedItems == null || selectedItems.Count == 0)
            {
                WpfMsg.Show("コピーするオブジェクトを選択してください。",
                    "情報", WpfMsgB.OK, WpfMsgI.Information);
                return;
            }
            if (selectedItems.Count > 1)
            {
                WpfMsg.Show("コピーは1件ずつ行ってください。",
                    "情報", WpfMsgB.OK, WpfMsgI.Information);
                return;
            }

            var item = (ObjectInfo)selectedItems[0]!;

            if (item.ObjectType != "TABLE")
            {
                WpfMsg.Show("コピーはテーブルのみ対応しています。",
                    "情報", WpfMsgB.OK, WpfMsgI.Information);
                return;
            }

            var existingNames = _listTab.Objects
                .Select(o => o.Name)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var dialog = new Views.CopyDialog(item.Name, existingNames)
            {
                Owner = WpfApp.Current.MainWindow
            };
            if (dialog.ShowDialog() != true) return;

            try
            {
                _schemaService.CopyTable(item.Name, dialog.DestName);
                LoadObjectList(CurrentCategory);
            }
            catch (Exception ex)
            {
                WpfMsg.Show($"コピーエラー: {ex.Message}", "エラー", WpfMsgB.OK, WpfMsgI.Error);
            }
        }

        // ── テーブル詳細ウィンドウを開く ─────────────────────────────────────

        [RelayCommand]
        private void OpenDetail(ObjectInfo? item)
        {
            if (item == null || _schemaService == null) return;

            if (item.IsOpen)
            {
                foreach (WpfWin w in WpfApp.Current.Windows)
                {
                    if (w is TableDetailWindow tdw
                     && (tdw.DataContext as TableDetailViewModel)?.TableName == item.Name)
                    {
                        tdw.Activate();
                        return;
                    }
                }
            }

            var vm  = new TableDetailViewModel(item.Name, item.ObjectType, _schemaService);
            var win = new TableDetailWindow(vm, item);
            win.Show();
        }

        // ── クエリタブ管琁E────────────────────────────────────────────────────

        private void AddNewQueryTab()
        {
            _addingTab = true;
            _queryTabCount++;
            var newTab = new QueryTab(_queryTabCount, _dbService, _schemaService, _appSettings);
            Tabs.Insert(Tabs.Count - 1, newTab);   // + の手前に挿入
            SetProperty(ref _selectedTab, newTab, nameof(SelectedTab));
            _addingTab = false;
        }

        [RelayCommand]
        private void CloseTab(TabBase? tab)
        {
            if (tab == null || !tab.IsClosable) return;
            int idx = Tabs.IndexOf(tab);
            if (idx > 0) SelectedTab = Tabs[idx - 1];
            Tabs.Remove(tab);
        }

        // ── スクリプト出劁E────────────────────────────────────────────────────

        [RelayCommand]
        private void OpenScriptDialog()
        {
            if (_scriptService == null || _schemaService == null)
            {
                WpfMsg.Show("先に接続してください", "情報", WpfMsgB.OK, WpfMsgI.Information);
                return;
            }
            var vm     = new ScriptViewModel(_scriptService, _schemaService, _appSettings);
            var dialog = new ScriptDialog(vm) { Owner = WpfApp.Current.MainWindow };
            dialog.ShowDialog();
        }

        // ── テーブル定義書 Excel 出力（メニューから） ────────────────────────

        [RelayCommand]
        private void ExportTableDefinitionMenu()
        {
            if (_schemaService == null)
            {
                WpfMsg.Show("先に接続してください。", "情報", WpfMsgB.OK, WpfMsgI.Information);
                return;
            }
            var vm     = new ExportDefinitionViewModel(_schemaService, _appSettings);
            var dialog = new Views.ExportDefinitionDialog(vm) { Owner = WpfApp.Current.MainWindow };
            dialog.ShowDialog();
        }

        // ── テーブル定義書 Excel 出力 ─────────────────────────────────────────

        [RelayCommand]
        private void ExportTableDefinition(IList<ObjectInfo>? selectedItems)
        {
            if (_schemaService == null || selectedItems == null || selectedItems.Count == 0)
            {
                WpfMsg.Show("テーブルを選択してください", "情報", WpfMsgB.OK, WpfMsgI.Information);
                return;
            }

            var initDir = string.IsNullOrEmpty(_appSettings.LastScriptFolder)
                ? (_appSettings.WorkFolder.Length > 0 ? _appSettings.WorkFolder
                   : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments))
                : _appSettings.LastScriptFolder;

            var dialog = new WinSave
            {
                Filter           = "Excel ファイル (*.xlsx)|*.xlsx",
                FileName         = "テーブル定義書.xlsx",
                DefaultExt       = "xlsx",
                InitialDirectory = initDir
            };
            if (dialog.ShowDialog() != true) return;

            try
            {
                using var wb = new XLWorkbook();
                foreach (var item in selectedItems)
                {
                    var cols = _schemaService.GetColumns(item.Name);
                    var ws   = wb.AddWorksheet(item.Name.Length > 31 ? item.Name[..31] : item.Name);

                    var headers = new[] { "列名", "論理名", "データ型", "NULL", "PK", "FK", "デフォルト" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        ws.Cell(1, i + 1).Value = headers[i];
                        ws.Cell(1, i + 1).Style.Font.Bold = true;
                        ws.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.LightSteelBlue;
                    }

                    for (int r = 0; r < cols.Count; r++)
                    {
                        var col = cols[r];
                        ws.Cell(r + 2, 1).Value = col.ColumnName;
                        ws.Cell(r + 2, 2).Value = col.LogicalName;
                        ws.Cell(r + 2, 3).Value = col.DataType;
                        ws.Cell(r + 2, 4).Value = col.NullableText;
                        ws.Cell(r + 2, 5).Value = col.PrimaryKeyText;
                        ws.Cell(r + 2, 6).Value = col.ForeignKeyText;
                        ws.Cell(r + 2, 7).Value = col.DefaultValue;
                    }
                    ws.Columns().AdjustToContents();
                }

                wb.SaveAs(dialog.FileName);
                _appSettings.LastScriptFolder = Path.GetDirectoryName(dialog.FileName) ?? _appSettings.LastScriptFolder;
                SettingsService.Save(_appSettings);
                WpfMsg.Show($"保存しました:\n{dialog.FileName}", "完了", WpfMsgB.OK, WpfMsgI.Information);
            }
            catch (Exception ex)
            {
                WpfMsg.Show($"エラー: {ex.Message}", "エラー", WpfMsgB.OK, WpfMsgI.Error);
            }
        }

        // ── SQL 実行履歴 ─────────────────────────────────────────────────────

        [RelayCommand]
        private void OpenSqlHistory()
        {
            var win = new Views.SqlHistoryWindow(_dbService) { Owner = WpfApp.Current.MainWindow };
            win.Show();
        }

        // ── データ移行 ────────────────────────────────────────────────────────

        [RelayCommand]
        private void OpenDataMigration()
        {
            if (!IsConnected)
            {
                WpfMsg.Show("先に接続してください。", "情報", WpfMsgB.OK, WpfMsgI.Information);
                return;
            }
            var migService = new Services.DataMigrationService(_dbService);
            var vm         = new DataMigrationViewModel(migService, _appSettings);
            var dialog     = new Views.DataMigrationWizard(vm) { Owner = WpfApp.Current.MainWindow };
            dialog.ShowDialog();
        }

        // ── バックアップ ──────────────────────────────────────────────────────

        [RelayCommand]
        private void Backup()
        {
            if (_backupService == null || string.IsNullOrEmpty(_dbService.DatabaseName))
            {
                WpfMsg.Show("先に接続してください", "情報", WpfMsgB.OK, WpfMsgI.Information);
                return;
            }
            var vm     = new BackupViewModel(_backupService, _dbService.DatabaseName!, _appSettings);
            var dialog = new BackupDialog(vm) { Owner = WpfApp.Current.MainWindow };
            dialog.ShowDialog();
        }

        // ── リストア ──────────────────────────────────────────────────────────

        [RelayCommand]
        private async Task RestoreAsync()
        {
            if (_backupService == null || string.IsNullOrEmpty(_dbService.DatabaseName))
            {
                WpfMsg.Show("先に接続してください", "情報", WpfMsgB.OK, WpfMsgI.Information);
                return;
            }

            var initPath = string.IsNullOrEmpty(_appSettings.WorkFolder)
                ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                : _appSettings.WorkFolder;

            var openDlg = new WinOpen
            {
                Filter           = "バックアップファイル (*.bak)|*.bak|すべてのファイル (*.*)|*.*",
                Title            = "リストアするバックアップファイルを選択",
                InitialDirectory = initPath
            };
            if (openDlg.ShowDialog() != true) return;

            var bakFile = openDlg.FileName;
            var dbName  = _dbService.DatabaseName!;

            var confirm = WpfMsg.Show(
                $"データベース [{dbName}] を以下のファイルからリストアします。\n\n{bakFile}\n\n既存データは上書きされます。続行しますか？",
                "リストア確認",
                WpfMsgB.YesNo, WpfMsgI.Warning);
            if (confirm != WpfMsgR.Yes) return;

            try
            {
                await Task.Run(() => _backupService.Restore(dbName, bakFile));
                WpfMsg.Show("リストア完了", "完了", WpfMsgB.OK, WpfMsgI.Information);
            }
            catch (Exception ex)
            {
                WpfMsg.Show($"リストアエラー: {ex.Message}", "エラー", WpfMsgB.OK, WpfMsgI.Error);
            }
        }
    }
}
