using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SqlServerTool.Models;
using SqlServerTool.Services;
using System.IO;
using System.Windows.Forms;
using WpfMsg  = System.Windows.MessageBox;
using WpfMsgB = System.Windows.MessageBoxButton;
using WpfMsgI = System.Windows.MessageBoxImage;
using WpfMsgR = System.Windows.MessageBoxResult;

namespace SqlServerTool.ViewModels
{
    public partial class BackupViewModel : ObservableObject
    {
        private readonly BackupService _backupService;
        private readonly string        _databaseName;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(FullPath))]
        private string outputFolder = string.Empty;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(FullPath))]
        private string fileName = string.Empty;

        [ObservableProperty] private bool   confirmOverwrite = true;
        [ObservableProperty] private bool   isBusy           = false;
        [ObservableProperty] private string statusMessage     = string.Empty;

        public string FullPath => (string.IsNullOrWhiteSpace(OutputFolder) || string.IsNullOrWhiteSpace(FileName))
            ? string.Empty
            : Path.Combine(OutputFolder, FileName);

        public event Action? RequestClose;

        public BackupViewModel(BackupService backupService, string databaseName, AppSettings settings)
        {
            _backupService = backupService;
            _databaseName  = databaseName;

            // 初期フォルダ：作業用フォルダ → MyDocuments
            OutputFolder = string.IsNullOrEmpty(settings.WorkFolder)
                ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                : settings.WorkFolder;

            // デフォルトファイル名：DB名_日付.bak
            FileName = $"{databaseName}_{DateTime.Now:yyyyMMdd}.bak";
        }

        [RelayCommand]
        private void BrowseFolder()
        {
            using var dialog = new FolderBrowserDialog
            {
                Description  = "バックアップ先フォルダを選択",
                SelectedPath = OutputFolder
            };
            if (dialog.ShowDialog() == DialogResult.OK)
                OutputFolder = dialog.SelectedPath;
        }

        [RelayCommand]
        private async Task ExecuteAsync()
        {
            if (string.IsNullOrWhiteSpace(OutputFolder))
            {
                StatusMessage = "出力フォルダを指定してください";
                return;
            }
            if (string.IsNullOrWhiteSpace(FileName))
            {
                StatusMessage = "ファイル名を指定してください";
                return;
            }

            var filePath = FullPath;

            // 上書き確認
            if (ConfirmOverwrite && File.Exists(filePath))
            {
                var result = WpfMsg.Show(
                    $"以下のファイルが既に存在します。上書きしますか？\n\n{filePath}",
                    "上書き確認",
                    WpfMsgB.YesNo, WpfMsgI.Warning);
                if (result != WpfMsgR.Yes) return;
            }

            IsBusy = true;
            StatusMessage = "バックアップ実行中...";

            try
            {
                await Task.Run(() => _backupService.BackupToFile(_databaseName, filePath));
                StatusMessage = "完了";
                WpfMsg.Show($"バックアップ完了:\n{filePath}", "完了",
                    WpfMsgB.OK, WpfMsgI.Information);
                RequestClose?.Invoke();
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

        [RelayCommand]
        private void Cancel() => RequestClose?.Invoke();
    }
}
