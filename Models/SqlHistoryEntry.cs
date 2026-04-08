namespace SqlServerTool.Models
{
    public class SqlHistoryEntry
    {
        public DateTime ExecutedAt    { get; set; }
        public string   OperationType { get; set; } = string.Empty;
        public string   ObjectName    { get; set; } = string.Empty;
        public string   Sql           { get; set; } = string.Empty;

        /// <summary>一覧表示用の短縮 SQL（最初の行）</summary>
        public string SqlPreview
        {
            get
            {
                var first = Sql.Split('\n').FirstOrDefault(l => !l.TrimStart().StartsWith("--") && l.Trim().Length > 0)
                            ?? Sql;
                return first.Length > 100 ? first[..100] + "…" : first;
            }
        }
    }

    /// <summary>SQL プレビューダイアログでキャンセルされた場合にスローする例外</summary>
    public class SqlPreviewCancelledException : Exception
    {
        public SqlPreviewCancelledException() : base("SQL実行がキャンセルされました") { }
    }
}
