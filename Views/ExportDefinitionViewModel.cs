using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SqlServerTool.Models;
using SqlServerTool.Services;
using System.Collections.ObjectModel;
using System.IO;
using WinSave = Microsoft.Win32.SaveFileDialog;
using WpfMsg  = System.Windows.MessageBox;
using WpfMsgB = System.Windows.MessageBoxButton;
using WpfMsgI = System.Windows.MessageBoxImage;

namespace SqlServerTool.ViewModels
{
    public partial class ExportDefinitionViewModel : ObservableObject
    {
        private readonly SchemaService _schemaService;
        private readonly AppSettings   _appSettings;

        [ObservableProperty] private string outputFilePath    = string.Empty;
        [ObservableProperty] private bool   overwriteExisting = true;
        [ObservableProperty] private bool   includeLogicalName = true;
        [ObservableProperty] private bool   colorHeaders       = true;
        [ObservableProperty] private string statusMessage      = string.Empty;

        public ObservableCollection<SelectableItem> Tables { get; } = new();

        public ExportDefinitionViewModel(SchemaService schemaService, AppSettings appSettings)
        {
            _schemaService = schemaService;
            _appSettings   = appSettings;

            var initDir = string.IsNullOrEmpty(appSettings.LastScriptFolder)
                ? (appSettings.WorkFolder.Length > 0 ? appSettings.WorkFolder
                   : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments))
                : appSettings.LastScriptFolder;

            OutputFilePath = Path.Combine(initDir, "テーブル定義書.xlsx");
            LoadTables();
        }

        private void LoadTables()
        {
            try
            {
                foreach (var name in _schemaService.GetTableNames())
                    Tables.Add(new SelectableItem(name, false));
            }
            catch (Exception ex)
            {
                StatusMessage = $"テーブル一覧取得エラー: {ex.Message}";
            }
        }

        [RelayCommand]
        private void SelectAll()
        {
            foreach (var t in Tables) t.IsSelected = true;
        }

        [RelayCommand]
        private void DeselectAll()
        {
            foreach (var t in Tables) t.IsSelected = false;
        }

        [RelayCommand]
        private void BrowseFile()
        {
            var dialog = new WinSave
            {
                Filter           = "Excel ファイル (*.xlsx)|*.xlsx",
                FileName         = Path.GetFileName(OutputFilePath),
                DefaultExt       = "xlsx",
                InitialDirectory = Path.GetDirectoryName(OutputFilePath)
                                   ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };
            if (dialog.ShowDialog() == true)
                OutputFilePath = dialog.FileName;
        }

        [RelayCommand]
        private void Start()
        {
            var selected = Tables.Where(t => t.IsSelected).Select(t => t.Name).ToList();
            if (selected.Count == 0)
            {
                StatusMessage = "エクスポートするテーブルを選択してください。";
                return;
            }
            if (string.IsNullOrWhiteSpace(OutputFilePath))
            {
                StatusMessage = "出力先ファイルを指定してください。";
                return;
            }
            if (File.Exists(OutputFilePath) && !OverwriteExisting)
            {
                StatusMessage = "同名ファイルが存在します。上書きを許可するか別のファイル名を指定してください。";
                return;
            }

            try
            {
                Export(selected);
                _appSettings.LastScriptFolder =
                    Path.GetDirectoryName(OutputFilePath) ?? _appSettings.LastScriptFolder;
                SettingsService.Save(_appSettings);

                StatusMessage = $"出力完了: {Path.GetFileName(OutputFilePath)}（{selected.Count} テーブル）";
                WpfMsg.Show($"出力しました:\n{OutputFilePath}", "完了", WpfMsgB.OK, WpfMsgI.Information);
            }
            catch (Exception ex)
            {
                StatusMessage = $"エラー: {ex.Message}";
                WpfMsg.Show($"出力エラー:\n{ex.Message}", "エラー", WpfMsgB.OK, WpfMsgI.Error);
            }
        }

        private void Export(IReadOnlyList<string> tableNames)
        {
            using var wb = new XLWorkbook();

            foreach (var name in tableNames)
            {
                var cols = _schemaService.GetColumns(name);
                var sheetName = name.Length > 31 ? name[..31] : name;
                var ws = wb.AddWorksheet(sheetName);

                // ヘッダー定義
                var headers = IncludeLogicalName
                    ? new[] { "列名", "論理名", "データ型", "NULL", "PK", "FK", "デフォルト" }
                    : new[] { "列名", "データ型", "NULL", "PK", "FK", "デフォルト" };

                for (int i = 0; i < headers.Length; i++)
                {
                    var cell = ws.Cell(1, i + 1);
                    cell.Value = headers[i];
                    cell.Style.Font.Bold = true;
                    if (ColorHeaders)
                        cell.Style.Fill.BackgroundColor = XLColor.LightSteelBlue;
                }

                for (int r = 0; r < cols.Count; r++)
                {
                    var col = cols[r];
                    int c = 1;
                    ws.Cell(r + 2, c++).Value = col.ColumnName;
                    if (IncludeLogicalName)
                        ws.Cell(r + 2, c++).Value = col.LogicalName;
                    ws.Cell(r + 2, c++).Value = col.DataType;
                    ws.Cell(r + 2, c++).Value = col.NullableText;
                    ws.Cell(r + 2, c++).Value = col.PrimaryKeyText;
                    ws.Cell(r + 2, c++).Value = col.ForeignKeyText;
                    ws.Cell(r + 2, c++).Value = col.DefaultValue;
                }

                ws.Columns().AdjustToContents();
            }

            wb.SaveAs(OutputFilePath);
        }
    }
}
