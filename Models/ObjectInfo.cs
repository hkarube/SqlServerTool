using CommunityToolkit.Mvvm.ComponentModel;

namespace SqlServerTool.Models
{
    public partial class ObjectInfo : ObservableObject
    {
        public string   Name       { get; set; } = string.Empty;
        public string   ObjectType { get; set; } = string.Empty;
        public string   Owner      { get; set; } = string.Empty;
        public long     RowCount   { get; set; }
        public DateTime? CreateDate { get; set; }
        public string   Comment    { get; set; } = string.Empty;

        [ObservableProperty]
        private bool isOpen = false;

        public string RowCountText   => ObjectType == "TABLE" ? RowCount.ToString("#,##0") : "-";
        public string CreateDateText => CreateDate?.ToString("yyyy/MM/dd HH:mm") ?? "";
    }
}
