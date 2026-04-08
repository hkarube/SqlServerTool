using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SqlServerTool.Models;
using SqlServerTool.Services;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows.Forms;

namespace SqlServerTool.ViewModels
{
    public partial class ScriptViewModel : ObservableObject
    {
        private readonly ScriptService  _scriptService;
        private readonly SchemaService  _schemaService;
        private readonly AppSettings    _appSettings;

        [ObservableProperty] private string outputFolder      = string.Empty;
        [ObservableProperty] private bool   overwriteExisting = true;
        [ObservableProperty] private bool   includeDrop        = true;
        [ObservableProperty] private bool   includeData        = true;
        [ObservableProperty] private bool   includeFilegroup   = true;
        [ObservableProperty] private bool   includeViewColumns = true;
        [ObservableProperty] private string statusMessage      = string.Empty;
        [ObservableProperty] private bool   isBusy             = false;

        public ObservableCollection<SelectableItem> Tables    { get; } = new();
        public ObservableCollection<SelectableItem> Views     { get; } = new();
        public ObservableCollection<SelectableItem> StoredProcs { get; } = new();
        public ObservableCollection<SelectableItem> Functions { get; } = new();

        public ScriptViewModel(ScriptService scriptService, SchemaService schemaService, AppSettings appSettings)
        {
            _scriptService = scriptService;
            _schemaService = schemaService;
            _appSettings   = appSettings;

            OutputFolder = string.IsNullOrEmpty(appSettings.LastScriptFolder)
                ? (appSettings.WorkFolder.Length > 0 ? appSettings.WorkFolder
                   : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments))
                : appSettings.LastScriptFolder;

            LoadObjects();
        }

        private void LoadObjects()
        {
            try
            {
                foreach (var t in _schemaService.GetTableNames())
                    Tables.Add(new SelectableItem(t, false));
                foreach (var v in _schemaService.GetViewNames())
                    Views.Add(new SelectableItem(v, false));
                foreach (var p in _schemaService.GetStoredProcNames())
                    StoredProcs.Add(new SelectableItem(p, false));
                foreach (var f in _schemaService.GetFunctionNames())
                    Functions.Add(new SelectableItem(f, false));
            }
            catch (Exception ex)
            {
                StatusMessage = $"一覧取得エラー: {ex.Message}";
            }
        }

        [RelayCommand]
        private void BrowseFolder()
        {
            using var dialog = new FolderBrowserDialog { SelectedPath = OutputFolder };
            if (dialog.ShowDialog() == DialogResult.OK)
                OutputFolder = dialog.SelectedPath;
        }

        [RelayCommand]
        private async Task StartAsync()
        {
            if (string.IsNullOrWhiteSpace(OutputFolder))
            {
                StatusMessage = "出力先を指定してください";
                return;
            }

            IsBusy = true;
            StatusMessage = "処理中...";

            var options = new ScriptOptions
            {
                OutputFolder       = OutputFolder,
                OverwriteExisting  = OverwriteExisting,
                IncludeDrop        = IncludeDrop,
                IncludeData        = IncludeData,
                IncludeFilegroup   = IncludeFilegroup,
                IncludeViewColumns = IncludeViewColumns,
            };

            var selTables = Tables    .Where(x => x.IsSelected).Select(x => x.Name).ToList();
            var selViews  = Views     .Where(x => x.IsSelected).Select(x => x.Name).ToList();
            var selProcs  = StoredProcs.Where(x => x.IsSelected).Select(x => x.Name).ToList();
            var selFuncs  = Functions .Where(x => x.IsSelected).Select(x => x.Name).ToList();

            var progress = new Progress<string>(msg => StatusMessage = msg);

            try
            {
                await Task.Run(() =>
                {
                    _scriptService.GenerateTableScripts(selTables, options, progress);
                    _scriptService.GenerateViewScripts(selViews, options, progress);
                    _scriptService.GenerateCodeObjectScripts(selProcs,  "PROCEDURE", options, progress);
                    _scriptService.GenerateCodeObjectScripts(selFuncs,  "FUNCTION",  options, progress);
                });

                _appSettings.LastScriptFolder = OutputFolder;
                SettingsService.Save(_appSettings);
                StatusMessage = $"完了。出力先: {OutputFolder}";
            }
            catch (Exception ex)
            {
                StatusMessage = $"エラー: {ex.Message}";
            }
            finally
            {
                IsBusy = false;
            }
        }

        [RelayCommand] private void SelectAll()
        {
            foreach (var x in Tables.Concat(Views).Concat(StoredProcs).Concat(Functions))
                x.IsSelected = true;
        }

        [RelayCommand] private void DeselectAll()
        {
            foreach (var x in Tables.Concat(Views).Concat(StoredProcs).Concat(Functions))
                x.IsSelected = false;
        }
    }

    public partial class SelectableItem : ObservableObject
    {
        public string Name { get; }
        [ObservableProperty] private bool isSelected;
        public SelectableItem(string name, bool sel = false) { Name = name; IsSelected = sel; }
    }
}
