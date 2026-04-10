using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SqlServerTool.Services;
using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using WinOpen = Microsoft.Win32.OpenFileDialog;
using WinSave = Microsoft.Win32.SaveFileDialog;
using WpfMsg  = System.Windows.MessageBox;
using WpfMsgB = System.Windows.MessageBoxButton;
using WpfMsgI = System.Windows.MessageBoxImage;

namespace SqlServerTool.ViewModels
{
    public partial class BulkCopyViewModel : ObservableObject
    {
        private readonly BulkCopyService _service;
        private readonly string _initFolder;

        // ── エクスポート ──────────────────────────────────────────────────────
        [ObservableProperty] private string exportFolder      = string.Empty;
        [ObservableProperty] private string exportEncoding    = "UTF-8";
        [ObservableProperty] private string exportDelimiter   = "カンマ ( , )";
        [ObservableProperty] private bool   exportHeader      = true;
        [ObservableProperty] private string exportStatus      = string.Empty;

        public ObservableCollection<SelectableItem> ExportTables { get; } = new();

        // ── インポート ──────────────────────────────────────────────────────
        [ObservableProperty] private string importFilePath    = string.Empty;
        [ObservableProperty] private string importEncoding    = "UTF-8";
        [ObservableProperty] private string importDelimiter   = "カンマ ( , )";
        [ObservableProperty] private bool   importHeader      = true;
        [ObservableProperty] private bool   truncateFirst     = false;
        [ObservableProperty] private string importTargetTable = string.Empty;
        [ObservableProperty] private string importStatus      = string.Empty;

        public ObservableCollection<string> AllTableNames { get; } = new();

        public IReadOnlyList<string> EncodingOptions { get; } =
            new[] { "UTF-8", "Shift-JIS" };

        public IReadOnlyList<string> DelimiterOptions { get; } =
            new[] { "カンマ ( , )", "タブ" };

        public BulkCopyViewModel(BulkCopyService service, string initFolder)
        {
            _service    = service;
            _initFolder = initFolder;
            ExportFolder = initFolder;
            LoadTables();
        }

        private void LoadTables()
        {
            try
            {
                foreach (var name in _service.GetTableNames())
                {
                    ExportTables.Add(new SelectableItem(name, false));
                    AllTableNames.Add(name);
                }
            }
            catch (Exception ex)
            {
                ExportStatus = $"テーブル一覧取得エラー: {ex.Message}";
            }
        }

        // ── エクスポート コマンド ─────────────────────────────────────────────

        [RelayCommand]
        private void ExportSelectAll()   { foreach (var t in ExportTables) t.IsSelected = true; }

        [RelayCommand]
        private void ExportDeselectAll() { foreach (var t in ExportTables) t.IsSelected = false; }

        [RelayCommand]
        private void BrowseExportFolder()
        {
            using var dlg = new System.Windows.Forms.FolderBrowserDialog
            {
                SelectedPath = ExportFolder
            };
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                ExportFolder = dlg.SelectedPath;
        }

        [RelayCommand]
        private async Task StartExport()
        {
            var targets = ExportTables.Where(t => t.IsSelected).Select(t => t.Name).ToList();
            if (targets.Count == 0) { ExportStatus = "エクスポートするテーブルを選択してください。"; return; }
            if (string.IsNullOrWhiteSpace(ExportFolder)) { ExportStatus = "出力先フォルダを指定してください。"; return; }
            if (!Directory.Exists(ExportFolder)) { ExportStatus = "指定フォルダが存在しません。"; return; }

            var enc   = GetEncoding(ExportEncoding);
            var delim = GetDelimiter(ExportDelimiter);

            ExportStatus = "エクスポート中...";
            int totalRows = 0;
            var errors = new List<string>();

            await Task.Run(() =>
            {
                foreach (var table in targets)
                {
                    try
                    {
                        var path = Path.Combine(ExportFolder, $"{table}.csv");
                        int rows = _service.ExportToCsv(table, path, enc, delim, ExportHeader);
                        totalRows += rows;
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"{table}: {ex.Message}");
                    }
                }
            });

            if (errors.Count > 0)
                ExportStatus = $"完了（エラーあり）:\n{string.Join("\n", errors)}";
            else
                ExportStatus = $"エクスポート完了: {targets.Count} テーブル / {totalRows:#,##0} 件";
        }

        // ── インポート コマンド ─────────────────────────────────────────────

        [RelayCommand]
        private void BrowseImportFile()
        {
            var dlg = new WinOpen
            {
                Filter           = "CSV ファイル (*.csv)|*.csv|すべてのファイル (*.*)|*.*",
                InitialDirectory = _initFolder
            };
            if (dlg.ShowDialog() == true)
            {
                ImportFilePath = dlg.FileName;
                // ファイル名からテーブル名を自動推測
                var guess = Path.GetFileNameWithoutExtension(dlg.FileName);
                if (AllTableNames.Contains(guess, StringComparer.OrdinalIgnoreCase))
                    ImportTargetTable = AllTableNames.First(
                        n => string.Equals(n, guess, StringComparison.OrdinalIgnoreCase));
            }
        }

        [RelayCommand]
        private async Task StartImport()
        {
            if (string.IsNullOrWhiteSpace(ImportFilePath) || !File.Exists(ImportFilePath))
            { ImportStatus = "CSVファイルを選択してください。"; return; }
            if (string.IsNullOrWhiteSpace(ImportTargetTable))
            { ImportStatus = "インポート先テーブルを選択してください。"; return; }

            var enc   = GetEncoding(ImportEncoding);
            var delim = GetDelimiter(ImportDelimiter);

            var truncMsg = TruncateFirst ? "既存データを削除してからインポートします。" : "既存データに追記します。";
            var confirm = WpfMsg.Show(
                $"[{ImportTargetTable}] に CSVをインポートします。\n{truncMsg}\n\n続行しますか？",
                "インポート確認", WpfMsgB.YesNo, WpfMsgI.Question);
            if (confirm != System.Windows.MessageBoxResult.Yes) return;

            ImportStatus = "インポート中...";
            try
            {
                int rows = 0;
                await Task.Run(() =>
                    rows = _service.ImportFromCsv(
                        ImportTargetTable, ImportFilePath,
                        enc, delim, ImportHeader, TruncateFirst));

                ImportStatus = $"インポート完了: {rows:#,##0} 件";
            }
            catch (Exception ex)
            {
                ImportStatus = $"エラー: {ex.Message}";
                WpfMsg.Show($"インポートエラー:\n{ex.Message}", "エラー", WpfMsgB.OK, WpfMsgI.Error);
            }
        }

        // ── ヘルパー ──────────────────────────────────────────────────────────

        private static Encoding GetEncoding(string label) => label switch
        {
            "Shift-JIS" => Encoding.GetEncoding("shift_jis"),
            _           => new UTF8Encoding(encoderShouldEmitUTF8Identifier: true)
        };

        private static char GetDelimiter(string label) => label switch
        {
            "タブ" => '\t',
            _     => ','
        };
    }
}
