using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SqlServerTool.Models;
using SqlServerTool.Services;

namespace SqlServerTool.ViewModels
{
    public partial class ConnectionViewModel : ObservableObject
    {
        private readonly DatabaseService _dbService;

        [ObservableProperty] private string serverName    = string.Empty;
        [ObservableProperty] private string databaseName  = string.Empty;
        [ObservableProperty] private string userId        = "sa";
        [ObservableProperty] private string password      = string.Empty;
        [ObservableProperty] private string statusMessage = string.Empty;
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(IsNotWindowsAuth))]
        private bool useWindowsAuth = false;

        public bool IsNotWindowsAuth => !UseWindowsAuth;

        public bool IsConnected { get; private set; } = false;

        /// <summary>接続成功時に発火 → ダイアログ側で Close() を呼ぶ</summary>
        public event Action? RequestClose;

        public ConnectionViewModel(DatabaseService dbService)
        {
            _dbService = dbService;

            // 保存済み設定を起動時に読み込む
            var saved = SettingsService.Load();
            ServerName      = saved.ServerName;
            DatabaseName    = saved.DatabaseName;
            UserId          = string.IsNullOrEmpty(saved.UserId) ? "sa" : saved.UserId;
            Password        = saved.Password;
            UseWindowsAuth  = saved.UseWindowsAuth;
        }

        [RelayCommand]
        private void TestConnection()
        {
            var settings = BuildSettings();
            StatusMessage = _dbService.TestConnection(settings, out string error)
                ? "✅ 接続成功"
                : $"❌ 接続失敗：{error}";
        }

        [RelayCommand]
        private void Connect()
        {
            var settings = BuildSettings();
            if (_dbService.TestConnection(settings, out string error))
            {
                _dbService.SetConnection(settings);
                IsConnected = true;
                StatusMessage = "接続しました";

                // 設定を保存
                SettingsService.Save(new AppSettings
                {
                    ServerName     = ServerName,
                    DatabaseName   = DatabaseName,
                    UserId         = UserId,
                    Password       = Password,
                    UseWindowsAuth = UseWindowsAuth
                });

                RequestClose?.Invoke();
            }
            else
            {
                StatusMessage = $"❌ {error}";
            }
        }

        private ConnectionSettings BuildSettings() => new()
        {
            ServerName     = ServerName,
            DatabaseName   = DatabaseName,
            UserId         = UserId,
            Password       = Password,
            UseWindowsAuth = UseWindowsAuth
        };
    }
}
