using Microsoft.Data.SqlClient;
using SqlServerTool.Models;

namespace SqlServerTool.Services
{
    public class DatabaseService
    {
        private string _connectionString = string.Empty;
        private ConnectionSettings? _settings;

        public string? DatabaseName => _settings?.DatabaseName;
        public string? ServerName   => _settings?.ServerName;
        public bool    IsConnected  => _settings != null && !string.IsNullOrEmpty(_connectionString);

        public void SetConnection(ConnectionSettings settings)
        {
            _settings = settings;
            _connectionString = settings.BuildConnectionString();
        }

        public bool TestConnection(ConnectionSettings settings, out string errorMessage)
        {
            errorMessage = string.Empty;
            try
            {
                using var conn = new SqlConnection(settings.BuildConnectionString());
                conn.Open();
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        public SqlConnection GetConnection()
        {
            return new SqlConnection(_connectionString);
        }

        public SqlConnection GetMasterConnection()
        {
            if (_settings == null) throw new InvalidOperationException("未接続です");
            var masterSettings = new ConnectionSettings
            {
                ServerName     = _settings.ServerName,
                DatabaseName   = "master",
                UserId         = _settings.UserId,
                Password       = _settings.Password,
                UseWindowsAuth = _settings.UseWindowsAuth
            };
            return new SqlConnection(masterSettings.BuildConnectionString());
        }
    }
}
