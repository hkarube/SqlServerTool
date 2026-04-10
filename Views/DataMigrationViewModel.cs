using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Data.SqlClient;
using SqlServerTool.Models;
using SqlServerTool.Services;
using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using WpfMsg  = System.Windows.MessageBox;
using WpfMsgB = System.Windows.MessageBoxButton;
using WpfMsgI = System.Windows.MessageBoxImage;

namespace SqlServerTool.ViewModels
{
    public partial class DataMigrationViewModel : ObservableObject
    {
        private readonly DataMigrationService _service;
        private readonly AppSettings          _appSettings;

        // ── Step ナビゲーション ───────────────────────────────────────────────
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(IsStep1))]
        [NotifyPropertyChangedFor(nameof(IsStep2))]
        [NotifyPropertyChangedFor(nameof(IsStep3))]
        private int currentStep = 0;   // 0=Step1, 1=Step2, 2=Step3

        public bool IsStep1 => CurrentStep == 0;
        public bool IsStep2 => CurrentStep == 1;
        public bool IsStep3 => CurrentStep == 2;

        // ── Step1：テーブル選択・出力方式 ────────────────────────────────────
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(IsDbMode))]
        [NotifyPropertyChangedFor(nameof(IsCsvMode))]
        [NotifyPropertyChangedFor(nameof(IsCsvImportMode))]
        [NotifyPropertyChangedFor(nameof(IsNotCsvImportMode))]
        [NotifyPropertyChangedFor(nameof(IsDbOrCsvImport))]
        [NotifyPropertyChangedFor(nameof(ShowDbDestination))]
        [NotifyPropertyChangedFor(nameof(ShowExistingDbSettings))]
        [NotifyPropertyChangedFor(nameof(IsDbModeAndNew))]
        [NotifyPropertyChangedFor(nameof(IsCsvImportOrCsvToDb))]
        private string outputMethod = "DB";   // "DB" / "CSV" / "CSVImport"

        public ObservableCollection<MigrationTableItem> Tables { get; } = new();
        public bool IsDbMode           => OutputMethod == "DB";
        public bool IsCsvMode          => OutputMethod == "CSV";
        public bool IsCsvImportMode    => OutputMethod == "CSVImport";
        public bool IsNotCsvImportMode => OutputMethod != "CSVImport";
        /// <summary>コピー先 Existing/New ラジオを表示するモード（DB / CSVImport）</summary>
        public bool IsDbOrCsvImport    => OutputMethod == "DB" || OutputMethod == "CSVImport";

        // ── Step2：宛先 DB 設定 ───────────────────────────────────────────────
        // 既存 DB / 新規 DB
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(IsExistingDb))]
        [NotifyPropertyChangedFor(nameof(IsNewDb))]
        [NotifyPropertyChangedFor(nameof(ShowExistingDbSettings))]
        [NotifyPropertyChangedFor(nameof(IsDbModeAndNew))]
        private string destDbMode = "Existing";   // "Existing" / "New"
        public bool IsExistingDb => DestDbMode == "Existing";
        public bool IsNewDb      => DestDbMode == "New";

        // 既存 DB 接続
        [ObservableProperty] private string destServer      = string.Empty;
        [ObservableProperty] private bool   destWindowsAuth = true;
        [ObservableProperty] private string destUserId      = string.Empty;
        public string DestPassword { get; set; } = string.Empty;
        [ObservableProperty] private string destConnStatus  = string.Empty;
        public ObservableCollection<string> DestDatabases { get; } = new();
        [ObservableProperty] private string destDatabase    = string.Empty;

        // 新規 DB 接続
        [ObservableProperty] private string newDbServer       = string.Empty;
        [ObservableProperty] private string newDbUserId       = string.Empty;
        public string NewDbPassword { get; set; } = string.Empty;
        [ObservableProperty] private string newDbName         = string.Empty;
        [ObservableProperty] private string newDbDataFilePath = string.Empty;

        // オプション（DB コピー用）
        [ObservableProperty] private bool copyStructure    = true;
        [ObservableProperty] private bool copyData         = true;
        [ObservableProperty] private bool preserveIdentity = true;
        [ObservableProperty] private bool disableFk        = true;

        // ── Step2：CSV 出力設定 ───────────────────────────────────────────────
        [ObservableProperty] private string csvOutputFolder  = string.Empty;
        [ObservableProperty] private string csvEncoding      = "UTF-8";
        [ObservableProperty] private string csvDelimiter     = "カンマ ( , )";
        [ObservableProperty] private bool   csvIncludeHeader = true;

        public IReadOnlyList<string> EncodingOptions  { get; } = new[] { "UTF-8", "Shift-JIS" };
        public IReadOnlyList<string> DelimiterOptions { get; } = new[] { "カンマ ( , )", "タブ" };

        // ── Step2：CSV→DB / CSVImport 設定 ───────────────────────────────────
        /// <summary>CSV エクスポート後に宛先 DB へもインポートするか（CSV モードのみ）</summary>
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(ShowDbDestination))]
        [NotifyPropertyChangedFor(nameof(ShowExistingDbSettings))]
        [NotifyPropertyChangedFor(nameof(IsCsvImportOrCsvToDb))]
        private bool copyCsvToDb = false;

        /// <summary>インポート前にテーブルをトランケートするか</summary>
        [ObservableProperty] private bool truncateBeforeImport = false;

        /// <summary>既存 CSV ファイルのインポート元フォルダ（CSVImport モード）</summary>
        [ObservableProperty] private string csvImportFolder = string.Empty;

        // 可視性制御用の計算プロパティ
        public bool ShowDbDestination     => IsDbMode || (IsCsvMode && CopyCsvToDb) || IsCsvImportMode;
        public bool ShowExistingDbSettings =>
            (IsDbMode && IsExistingDb) || (IsCsvMode && CopyCsvToDb) || (IsCsvImportMode && IsExistingDb);
        /// <summary>新規DB 作成設定を表示（DB または CSVImport で「新規」選択時）</summary>
        public bool IsDbModeAndNew        => (IsDbMode || IsCsvImportMode) && IsNewDb;
        public bool IsCsvImportOrCsvToDb  => IsCsvImportMode || (IsCsvMode && CopyCsvToDb);

        // ── Step3：プレビュー・実行 ───────────────────────────────────────────
        [ObservableProperty] private string previewDestination = string.Empty;
        [ObservableProperty] private string cycleWarning       = string.Empty;
        [ObservableProperty] private bool   hasCycleWarning    = false;
        [ObservableProperty] private double progress           = 0.0;
        [ObservableProperty] private bool   isRunning          = false;
        [ObservableProperty] private bool   hasFailedTables    = false;

        public ObservableCollection<MigrationTableItem> PreviewTables { get; } = new();
        public ObservableCollection<string>             ExecutionLog  { get; } = new();

        // ────────────────────────────────────────────────────────────────────

        public DataMigrationViewModel(DataMigrationService service, AppSettings appSettings)
        {
            _service     = service;
            _appSettings = appSettings;

            CsvOutputFolder = string.IsNullOrEmpty(appSettings.WorkFolder)
                ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                : appSettings.WorkFolder;

            // 要件5: 新規 DB のファイル保存先にソース DB と同じフォルダを既定値として設定
            try { NewDbDataFilePath = _service.GetSourceDbDataFolder(); }
            catch { /* 取得失敗時は空のまま */ }

            LoadTables();
        }

        private void LoadTables()
        {
            try
            {
                Tables.Clear();
                foreach (var item in _service.GetSourceTables())
                    Tables.Add(item);
            }
            catch (Exception ex)
            {
                WpfMsg.Show($"テーブル一覧取得エラー:\n{ex.Message}", "エラー", WpfMsgB.OK, WpfMsgI.Error);
            }
        }

        // ── テーブル全選択 / 全解除 ───────────────────────────────────────────
        [RelayCommand] private void SelectAll()   { foreach (var t in Tables) t.IsSelected = true; }
        [RelayCommand] private void DeselectAll() { foreach (var t in Tables) t.IsSelected = false; }

        // ── Step2：接続テスト ─────────────────────────────────────────────────
        [RelayCommand]
        private void TestDestConnection()
        {
            var connStr = DataMigrationService.BuildConnectionString(
                DestServer, DestWindowsAuth, DestUserId, DestPassword);
            DestConnStatus = DataMigrationService.TestConnection(connStr, out var err)
                ? "接続成功"
                : $"接続失敗: {err}";
        }

        [RelayCommand]
        private void LoadDestDatabases()
        {
            try
            {
                var connStr = DataMigrationService.BuildConnectionString(
                    DestServer, DestWindowsAuth, DestUserId, DestPassword);
                DestDatabases.Clear();
                foreach (var db in DataMigrationService.GetDatabaseNames(connStr))
                    DestDatabases.Add(db);
                DestConnStatus = $"DB一覧取得完了（{DestDatabases.Count}件）";
            }
            catch (Exception ex)
            {
                DestConnStatus = $"エラー: {ex.Message}";
            }
        }

        // ── CSV 出力先フォルダ参照 ────────────────────────────────────────────
        [RelayCommand]
        private void BrowseCsvFolder()
        {
            using var dlg = new System.Windows.Forms.FolderBrowserDialog
            {
                SelectedPath = CsvOutputFolder
            };
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                CsvOutputFolder = dlg.SelectedPath;
        }

        // ── CSV インポート元フォルダ参照 ──────────────────────────────────────
        [RelayCommand]
        private void BrowseCsvImportFolder()
        {
            using var dlg = new System.Windows.Forms.FolderBrowserDialog
            {
                SelectedPath = CsvImportFolder
            };
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                CsvImportFolder = dlg.SelectedPath;
        }

        // ── 新規 DB データファイルパス参照 ───────────────────────────────────
        [RelayCommand]
        private void BrowseDataFilePath()
        {
            using var dlg = new System.Windows.Forms.FolderBrowserDialog
            {
                SelectedPath = NewDbDataFilePath
            };
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                NewDbDataFilePath = dlg.SelectedPath;
        }

        // ── テーブルの手動並び替え ────────────────────────────────────────────
        [RelayCommand]
        private void MoveUp(MigrationTableItem item)
        {
            int idx = PreviewTables.IndexOf(item);
            if (idx > 0) PreviewTables.Move(idx, idx - 1);
        }

        [RelayCommand]
        private void MoveDown(MigrationTableItem item)
        {
            int idx = PreviewTables.IndexOf(item);
            if (idx >= 0 && idx < PreviewTables.Count - 1) PreviewTables.Move(idx, idx + 1);
        }

        // ── ログをクリップボードにコピー ─────────────────────────────────────
        [RelayCommand]
        private void CopyLogToClipboard()
        {
            if (ExecutionLog.Count == 0) return;
            // ExecutionLog は最新が先頭なので逆順（時系列順）にして結合
            var text = string.Join(Environment.NewLine, ExecutionLog.Reverse());
            System.Windows.Clipboard.SetText(text);
        }

        // ── ナビゲーション ────────────────────────────────────────────────────
        [RelayCommand]
        private void NextStep()
        {
            if (CurrentStep == 0)
            {
                if (!IsCsvImportMode)
                {
                    // テーブルが1つも選択されていない場合は警告
                    if (!Tables.Any(t => t.IsSelected))
                    {
                        WpfMsg.Show("テーブルを1つ以上選択してください。",
                            "確認", WpfMsgB.OK, WpfMsgI.Information);
                        return;
                    }
                }
                // 全モード共通: Step2 へ（CSV モードでもスキップしない）
                CurrentStep = 1;
            }
            else if (CurrentStep == 1)
            {
                if (!ValidateStep2(out var err))
                {
                    WpfMsg.Show(err, "入力エラー", WpfMsgB.OK, WpfMsgI.Warning);
                    return;
                }
                BuildPreview(GetSelectedItems());
                CurrentStep = 2;
            }
        }

        /// <summary>
        /// CSV 出力のみ実行ボタン用：DB コピーなしで Step3 へ進む。
        /// </summary>
        [RelayCommand]
        private void ExportCsvOnly()
        {
            if (string.IsNullOrWhiteSpace(CsvOutputFolder))
            {
                WpfMsg.Show("CSV 出力フォルダを指定してください。",
                    "確認", WpfMsgB.OK, WpfMsgI.Information);
                return;
            }
            var selected = Tables.Where(t => t.IsSelected).ToList();
            if (selected.Count == 0)
            {
                WpfMsg.Show("テーブルを1つ以上選択してください。",
                    "確認", WpfMsgB.OK, WpfMsgI.Information);
                return;
            }
            CopyCsvToDb = false;   // DB コピーなし
            BuildPreview(selected);
            CurrentStep = 2;
        }

        [RelayCommand]
        private void PreviousStep()
        {
            if (CurrentStep == 2)
                CurrentStep = 1;
            else if (CurrentStep == 1)
                CurrentStep = 0;
        }

        private List<MigrationTableItem> GetSelectedItems()
        {
            if (IsCsvImportMode)
                return GetCsvImportItems();
            return Tables.Where(t => t.IsSelected).ToList();
        }

        /// <summary>インポート元フォルダの CSV ファイルを MigrationTableItem 一覧に変換する</summary>
        private List<MigrationTableItem> GetCsvImportItems()
        {
            if (!Directory.Exists(CsvImportFolder))
                return new List<MigrationTableItem>();
            return Directory.GetFiles(CsvImportFolder, "*.csv")
                .Select(f => new MigrationTableItem(Path.GetFileNameWithoutExtension(f), 0))
                .OrderBy(i => i.Name)
                .ToList();
        }

        private bool ValidateStep2(out string error)
        {
            error = string.Empty;

            if (IsCsvMode)
            {
                if (string.IsNullOrWhiteSpace(CsvOutputFolder))
                { error = "CSV 出力フォルダを指定してください。"; return false; }

                if (CopyCsvToDb)
                {
                    if (string.IsNullOrWhiteSpace(DestServer))
                    { error = "宛先サーバー名を入力してください。"; return false; }
                    if (string.IsNullOrWhiteSpace(DestDatabase))
                    { error = "宛先 DB を選択してください。"; return false; }
                    if (!DestWindowsAuth && string.IsNullOrWhiteSpace(DestUserId))
                    { error = "ユーザー ID を入力してください。"; return false; }
                }
                return true;
            }

            if (IsCsvImportMode)
            {
                if (string.IsNullOrWhiteSpace(CsvImportFolder))
                { error = "インポート元フォルダを指定してください。"; return false; }
                if (!Directory.Exists(CsvImportFolder))
                { error = "指定されたフォルダが存在しません。"; return false; }
                if (!Directory.GetFiles(CsvImportFolder, "*.csv").Any())
                { error = "フォルダ内に CSV ファイルが見つかりません。"; return false; }

                if (IsExistingDb)
                {
                    if (string.IsNullOrWhiteSpace(DestServer))
                    { error = "宛先サーバー名を入力してください。"; return false; }
                    if (string.IsNullOrWhiteSpace(DestDatabase))
                    { error = "宛先 DB を選択してください。"; return false; }
                    if (!DestWindowsAuth && string.IsNullOrWhiteSpace(DestUserId))
                    { error = "ユーザー ID を入力してください。"; return false; }
                    // DB 存在確認
                    if (!CheckDestDbExists(out error)) return false;
                }
                else
                {
                    // 新規 DB 作成
                    if (string.IsNullOrWhiteSpace(NewDbServer))
                    { error = "サーバー名を入力してください。"; return false; }
                    if (string.IsNullOrWhiteSpace(NewDbUserId))
                    { error = "ユーザー ID を入力してください。"; return false; }
                    if (string.IsNullOrWhiteSpace(NewDbPassword))
                    { error = "パスワードを入力してください。"; return false; }
                    if (string.IsNullOrWhiteSpace(NewDbName))
                    { error = "新規 DB 名を入力してください。"; return false; }
                }
                return true;
            }

            // DB モード
            if (IsExistingDb)
            {
                if (string.IsNullOrWhiteSpace(DestServer))
                { error = "宛先サーバー名を入力してください。"; return false; }
                if (string.IsNullOrWhiteSpace(DestDatabase))
                { error = "宛先 DB を選択してください。"; return false; }
                if (!DestWindowsAuth && string.IsNullOrWhiteSpace(DestUserId))
                { error = "ユーザー ID を入力してください。"; return false; }
                // DB 存在確認
                if (!CheckDestDbExists(out error)) return false;
            }
            else
            {
                if (string.IsNullOrWhiteSpace(NewDbServer))
                { error = "サーバー名を入力してください。"; return false; }
                if (string.IsNullOrWhiteSpace(NewDbUserId))
                { error = "新規 DB 作成用のユーザー ID を入力してください。"; return false; }
                if (string.IsNullOrWhiteSpace(NewDbPassword))
                { error = "新規 DB 作成用のパスワードを入力してください。"; return false; }
                if (string.IsNullOrWhiteSpace(NewDbName))
                { error = "新規 DB 名を入力してください。"; return false; }
            }
            if (!CopyStructure && !CopyData)
            { error = "「構造をコピー」か「データをコピー」を少なくとも1つ選択してください。"; return false; }
            return true;
        }

        private void BuildPreview(List<MigrationTableItem> selected)
        {
            PreviewTables.Clear();
            ExecutionLog.Clear();
            HasFailedTables = false;
            Progress        = 0;

            if (IsCsvImportMode)
            {
                // CSVImport モード: FK ソート不要
                HasCycleWarning = false;
                CycleWarning    = string.Empty;
                foreach (var item in selected)
                    PreviewTables.Add(new MigrationTableItem(item.Name, item.RowCount));
            }
            else
            {
                // DB / CSV モード: FK トポロジカルソート
                var names = selected.Select(t => t.Name).ToList();
                var (ordered, hasCycle) = _service.GetTopologicalOrder(names);

                HasCycleWarning = hasCycle;
                CycleWarning    = hasCycle
                    ? "循環参照が検出されました。外部キー制約を一時無効化するか、下の一覧で順序を手動調整してください。"
                    : string.Empty;

                foreach (var name in ordered)
                {
                    var src = selected.First(t =>
                        string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase));
                    PreviewTables.Add(new MigrationTableItem(src.Name, src.RowCount));
                }
            }

            // 宛先説明
            PreviewDestination = IsCsvImportMode && IsNewDb
                ? $"CSVインポート: {CsvImportFolder} → 新規DB [{NewDbName}] on {NewDbServer}"
                : IsCsvImportMode
                ? $"CSVインポート: {CsvImportFolder} → {DestServer}/[{DestDatabase}]"
                : IsCsvMode && CopyCsvToDb
                    ? $"CSV出力: {CsvOutputFolder}  次に DBインポート: {DestServer}/[{DestDatabase}]"
                    : IsCsvMode
                        ? $"CSV出力先: {CsvOutputFolder}"
                        : IsNewDb
                            ? $"新規DB作成: {NewDbServer} / [{NewDbName}]"
                            : $"既存DB: {DestServer} / [{DestDatabase}]";
        }

        // ── 実行 ─────────────────────────────────────────────────────────────
        [RelayCommand]
        private async Task StartExecution()
        {
            IsRunning       = true;
            HasFailedTables = false;
            Progress        = 0;
            ExecutionLog.Clear();
            foreach (var t in PreviewTables) { t.Status = ""; t.ErrorMessage = ""; }

            try
            {
                if (IsCsvImportMode)
                {
                    await ExecuteCsvImport();
                }
                else if (IsCsvMode)
                {
                    await ExecuteCsvExport();
                    if (CopyCsvToDb)
                        await ExecuteCsvImportToDb();
                }
                else if (IsNewDb)
                    await ExecuteNewDbCopy();
                else
                    await ExecuteExistingDbCopy();
            }
            finally
            {
                IsRunning = false;
            }
        }

        [RelayCommand]
        private async Task RetryFailed()
        {
            var failed = PreviewTables.Where(t => t.Status == "エラー").ToList();
            if (failed.Count == 0) return;
            foreach (var t in failed) { t.Status = ""; t.ErrorMessage = ""; }
            HasFailedTables = false;

            IsRunning = true;
            try
            {
                if (IsCsvImportMode)
                {
                    await ExecuteCsvImport(failed);
                }
                else if (IsCsvMode)
                {
                    await ExecuteCsvExport(failed);
                    if (CopyCsvToDb)
                        await ExecuteCsvImportToDb(failed);
                }
                else
                {
                    var connStr = GetDestConnectionString();
                    using var conn = new SqlConnection(connStr);
                    conn.Open();
                    await ExecuteDbCopy(conn, failed, isRetry: true);
                }
            }
            finally { IsRunning = false; }
        }

        // ── CSV エクスポート ──────────────────────────────────────────────────
        private async Task ExecuteCsvExport(List<MigrationTableItem>? targets = null)
        {
            targets ??= PreviewTables.ToList();
            var enc   = GetEncoding();
            var delim = GetDelimiter();
            int done  = 0;

            foreach (var item in targets)
            {
                item.Status = "実行中";
                AddLog($"[{DateTime.Now:HH:mm:ss}] {item.Name} エクスポート開始...");
                try
                {
                    var path = Path.Combine(CsvOutputFolder, $"{item.Name}.csv");
                    int rows = await Task.Run(() =>
                        _service.ExportToCsv(item.Name, path, enc, delim, CsvIncludeHeader));
                    item.Status = "完了";
                    AddLog($"[{DateTime.Now:HH:mm:ss}] {item.Name}: {rows:#,##0}件 → {Path.GetFileName(path)}");
                }
                catch (Exception ex)
                {
                    item.Status       = "エラー";
                    item.ErrorMessage = ex.Message;
                    HasFailedTables   = true;
                    AddLog($"[{DateTime.Now:HH:mm:ss}] {item.Name} エラー: {ex.Message}");
                }
                Progress = (double)++done / targets.Count;
            }
            AddLog(HasFailedTables
                ? "--- CSV完了（エラーあり）---"
                : "--- CSV エクスポートが完了しました ---");
        }

        // ── CSV → 宛先 DB インポート ──────────────────────────────────────────
        private async Task ExecuteCsvImportToDb(List<MigrationTableItem>? targets = null)
        {
            // CSV エクスポートが完了しているアイテムのみ対象
            targets ??= PreviewTables.Where(t => t.Status == "完了").ToList();
            if (targets.Count == 0)
            {
                AddLog($"[{DateTime.Now:HH:mm:ss}] DBインポート対象のテーブルがありません。");
                return;
            }

            var enc     = GetEncoding();
            var delim   = GetDelimiter();
            var connStr = GetDestConnectionString();
            using var destConn = new SqlConnection(connStr);
            destConn.Open();

            AddLog($"[{DateTime.Now:HH:mm:ss}] 宛先 DB へのインポートを開始します...");
            int done = 0;

            foreach (var item in targets)
            {
                item.Status = "DBインポート中";
                AddLog($"[{DateTime.Now:HH:mm:ss}] {item.Name} DBインポート中...");
                try
                {
                    var path = Path.Combine(CsvOutputFolder, $"{item.Name}.csv");
                    if (!File.Exists(path))
                        throw new FileNotFoundException($"CSV ファイルが見つかりません: {path}");

                    int rows = await Task.Run(() =>
                        DataMigrationService.ImportCsvToTable(
                            destConn, item.Name, path, enc, delim, CsvIncludeHeader, TruncateBeforeImport));
                    item.Status = "完了(DB)";
                    AddLog($"[{DateTime.Now:HH:mm:ss}] {item.Name}: {rows:#,##0}件 DBインポート完了");
                }
                catch (Exception ex)
                {
                    item.Status       = "エラー";
                    item.ErrorMessage = ex.Message;
                    HasFailedTables   = true;
                    AddLog($"[{DateTime.Now:HH:mm:ss}] {item.Name} DBインポートエラー: {ex.Message}");
                }
                Progress = (double)++done / targets.Count;
            }
            AddLog(HasFailedTables
                ? "--- DBインポート完了（エラーあり）---"
                : "--- DB へのインポートが完了しました ---");
        }

        // ── 既存 CSV → 宛先 DB インポート（CSVImport モード）────────────────
        private async Task ExecuteCsvImport(List<MigrationTableItem>? targets = null)
        {
            targets ??= PreviewTables.ToList();
            var enc   = GetEncoding();
            var delim = GetDelimiter();

            // 新規 DB 作成が必要な場合は先に作成する
            if (IsNewDb)
            {
                var serverConnStr = DataMigrationService.BuildConnectionString(
                    NewDbServer, false, NewDbUserId, NewDbPassword);
                AddLog($"[{DateTime.Now:HH:mm:ss}] データベース [{NewDbName}] を作成中...");
                try
                {
                    await Task.Run(() =>
                        DataMigrationService.CreateDatabase(serverConnStr, NewDbName, NewDbDataFilePath));
                    AddLog($"[{DateTime.Now:HH:mm:ss}] データベース [{NewDbName}] 作成完了");
                }
                catch (Exception ex)
                {
                    AddLog($"[{DateTime.Now:HH:mm:ss}] DB作成エラー: {ex.Message}");
                    WpfMsg.Show($"データベース作成に失敗しました。\n{ex.Message}", "エラー", WpfMsgB.OK, WpfMsgI.Error);
                    return;
                }
            }

            var connStr = IsNewDb
                ? DataMigrationService.BuildConnectionString(NewDbServer, false, NewDbUserId, NewDbPassword, NewDbName)
                : GetDestConnectionString();

            using var destConn = new SqlConnection(connStr);
            destConn.Open();

            int done = 0;
            foreach (var item in targets)
            {
                item.Status = "実行中";
                var csvPath = Path.Combine(CsvImportFolder, $"{item.Name}.csv");
                AddLog($"[{DateTime.Now:HH:mm:ss}] {item.Name} インポート開始...");
                try
                {
                    if (!File.Exists(csvPath))
                        throw new FileNotFoundException($"CSV ファイルが見つかりません: {csvPath}");

                    int rows = await Task.Run(() =>
                        DataMigrationService.ImportCsvToTable(
                            destConn, item.Name, csvPath, enc, delim, CsvIncludeHeader, TruncateBeforeImport));
                    item.Status = "完了";
                    AddLog($"[{DateTime.Now:HH:mm:ss}] {item.Name}: {rows:#,##0}件 完了");
                }
                catch (Exception ex)
                {
                    item.Status       = "エラー";
                    item.ErrorMessage = ex.Message;
                    HasFailedTables   = true;
                    AddLog($"[{DateTime.Now:HH:mm:ss}] {item.Name} エラー: {ex.Message}");
                }
                Progress = (double)++done / targets.Count;
            }
            AddLog(HasFailedTables
                ? "--- 完了（エラーあり）---"
                : "--- CSV インポートが完了しました ---");
        }

        // ── 既存 DB コピー ────────────────────────────────────────────────────
        private async Task ExecuteExistingDbCopy()
        {
            var connStr = GetDestConnectionString();
            using var conn = new SqlConnection(connStr);
            conn.Open();
            await ExecuteDbCopy(conn, PreviewTables.ToList());
        }

        // ── 新規 DB 作成 + コピー ─────────────────────────────────────────────
        private async Task ExecuteNewDbCopy()
        {
            var serverConnStr = DataMigrationService.BuildConnectionString(
                NewDbServer, false, NewDbUserId, NewDbPassword);

            AddLog($"[{DateTime.Now:HH:mm:ss}] データベース [{NewDbName}] を作成中...");
            try
            {
                await Task.Run(() =>
                    DataMigrationService.CreateDatabase(serverConnStr, NewDbName, NewDbDataFilePath));
                AddLog($"[{DateTime.Now:HH:mm:ss}] データベース [{NewDbName}] 作成完了");
            }
            catch (Exception ex)
            {
                AddLog($"[{DateTime.Now:HH:mm:ss}] DB作成エラー: {ex.Message}");
                WpfMsg.Show($"データベース作成に失敗しました。\n{ex.Message}", "エラー", WpfMsgB.OK, WpfMsgI.Error);
                return;
            }

            var dbConnStr = DataMigrationService.BuildConnectionString(
                NewDbServer, false, NewDbUserId, NewDbPassword, NewDbName);
            using var conn = new SqlConnection(dbConnStr);
            conn.Open();
            await ExecuteDbCopy(conn, PreviewTables.ToList());
        }

        // ── DB コピー共通処理 ─────────────────────────────────────────────────
        private async Task ExecuteDbCopy(
            SqlConnection destConn,
            List<MigrationTableItem> targets,
            bool isRetry = false)
        {
            if (DisableFk)
            {
                AddLog($"[{DateTime.Now:HH:mm:ss}] 外部キー制約を無効化中...");
                await Task.Run(() => DataMigrationService.DisableAllForeignKeys(destConn));
            }

            int done  = 0;
            int total = targets.Count;

            foreach (var item in targets)
            {
                item.Status = "実行中";

                if (CopyStructure)
                {
                    AddLog($"[{DateTime.Now:HH:mm:ss}] {item.Name} 構造作成中...");
                    try
                    {
                        var ddl = await Task.Run(() => _service.GetCreateTableScript(item.Name));
                        await Task.Run(() =>
                            new SqlCommand(ddl, destConn) { CommandTimeout = 120 }.ExecuteNonQuery());
                    }
                    catch (Exception ex)
                    {
                        item.Status       = "エラー";
                        item.ErrorMessage = ex.Message;
                        HasFailedTables   = true;
                        AddLog($"[{DateTime.Now:HH:mm:ss}] {item.Name} 構造エラー: {ex.Message}");
                        Progress = (double)++done / total;
                        continue;
                    }
                }

                if (CopyData)
                {
                    AddLog($"[{DateTime.Now:HH:mm:ss}] {item.Name} データコピー中...");
                    try
                    {
                        int rows = await Task.Run(() =>
                            _service.CopyTableData(destConn, item.Name, PreserveIdentity));
                        item.Status = "完了";
                        AddLog($"[{DateTime.Now:HH:mm:ss}] {item.Name}: {rows:#,##0}件 完了");
                    }
                    catch (Exception ex)
                    {
                        item.Status       = "エラー";
                        item.ErrorMessage = ex.Message;
                        HasFailedTables   = true;
                        AddLog($"[{DateTime.Now:HH:mm:ss}] {item.Name} エラー: {ex.Message}");
                    }
                }
                else
                {
                    item.Status = "完了";
                }
                Progress = (double)++done / total;
            }

            if (DisableFk)
            {
                AddLog($"[{DateTime.Now:HH:mm:ss}] 外部キー制約を有効化中...");
                try
                {
                    await Task.Run(() => DataMigrationService.EnableAllForeignKeys(destConn));
                }
                catch (Exception ex)
                {
                    AddLog($"[{DateTime.Now:HH:mm:ss}] FK有効化エラー: {ex.Message}");
                }
            }

            AddLog(HasFailedTables
                ? "--- 完了（エラーあり、「失敗分を再実行」で再試行できます）---"
                : "--- 全テーブルのコピーが完了しました ---");
        }

        // ── ヘルパー ──────────────────────────────────────────────────────────

        /// <summary>
        /// DestServer/DestDatabase が存在するか確認する。
        /// 失敗時は error にメッセージを設定して false を返す。
        /// </summary>
        private bool CheckDestDbExists(out string error)
        {
            error = string.Empty;
            try
            {
                var serverConnStr = DataMigrationService.BuildConnectionString(
                    DestServer, DestWindowsAuth, DestUserId, DestPassword);
                if (!DataMigrationService.DatabaseExists(serverConnStr, DestDatabase))
                {
                    error = $"データベース「{DestDatabase}」が見つかりません。\nDB 名を確認してください。";
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                error = $"DB 存在確認中にエラーが発生しました:\n{ex.Message}";
                return false;
            }
        }

        private string GetDestConnectionString() =>
            DataMigrationService.BuildConnectionString(
                DestServer, DestWindowsAuth, DestUserId, DestPassword, DestDatabase);

        private Encoding GetEncoding() =>
            CsvEncoding == "Shift-JIS"
                ? Encoding.GetEncoding("shift_jis")
                : new UTF8Encoding(true);

        private char GetDelimiter() =>
            CsvDelimiter == "タブ" ? '\t' : ',';

        private void AddLog(string message) =>
            System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                ExecutionLog.Insert(0, message));
    }
}
