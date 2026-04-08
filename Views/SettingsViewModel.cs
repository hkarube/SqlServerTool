using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SqlServerTool.Models;
using SqlServerTool.Services;
using System.Windows.Forms;

namespace SqlServerTool.ViewModels
{
    public partial class SettingsViewModel : ObservableObject
    {
        [ObservableProperty] private string serverName   = string.Empty;
        [ObservableProperty] private string databaseName = string.Empty;
        [ObservableProperty] private string userId       = "sa";
        [ObservableProperty] private string password     = string.Empty;
        [ObservableProperty] private string workFolder   = string.Empty;

        public event Action? RequestClose;

        public SettingsViewModel(AppSettings settings)
        {
            ServerName   = settings.ServerName;
            DatabaseName = settings.DatabaseName;
            UserId       = string.IsNullOrEmpty(settings.UserId) ? "sa" : settings.UserId;
            Password     = settings.Password;
            WorkFolder   = settings.WorkFolder;
        }

        [RelayCommand]
        private void BrowseWorkFolder()
        {
            using var dialog = new FolderBrowserDialog
            {
                Description  = "作業用フォルダを選択",
                SelectedPath = WorkFolder
            };
            if (dialog.ShowDialog() == DialogResult.OK)
                WorkFolder = dialog.SelectedPath;
        }

        [RelayCommand]
        private void Save()
        {
            // 既存設定を読み込み、接続情報と作業フォルダを更新
            var settings = SettingsService.Load();
            settings.ServerName   = ServerName;
            settings.DatabaseName = DatabaseName;
            settings.UserId       = UserId;
            settings.Password     = Password;
            settings.WorkFolder   = WorkFolder;
            SettingsService.Save(settings);
            RequestClose?.Invoke();
        }

        [RelayCommand]
        private void Cancel()
        {
            RequestClose?.Invoke();
        }
    }
}
