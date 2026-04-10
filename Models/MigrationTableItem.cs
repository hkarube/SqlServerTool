using CommunityToolkit.Mvvm.ComponentModel;

namespace SqlServerTool.Models
{
    public partial class MigrationTableItem : ObservableObject
    {
        public string Name     { get; }
        public long   RowCount { get; }

        [ObservableProperty] private bool   isSelected   = true;
        [ObservableProperty] private string status       = string.Empty;  // "" / "実行中" / "完了" / "エラー" / "スキップ"
        [ObservableProperty] private string errorMessage = string.Empty;

        public string RowCountText => RowCount >= 0 ? $"{RowCount:#,##0}" : "-";

        public MigrationTableItem(string name, long rowCount = -1)
        {
            Name     = name;
            RowCount = rowCount;
        }
    }
}
