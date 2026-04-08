namespace SqlServerTool.Models
{
    public class AppSettings
    {
        public string ServerName { get; set; } = string.Empty;
        public string DatabaseName { get; set; } = string.Empty;
        public string UserId         { get; set; } = "sa";
        public string Password       { get; set; } = string.Empty;
        public bool   UseWindowsAuth { get; set; } = false;
        /// <summary>バックアップ・リストア等で使う作業フォルダ</summary>
        public string WorkFolder { get; set; } = string.Empty;
        public string LastScriptFolder { get; set; } = string.Empty;
    }
}
