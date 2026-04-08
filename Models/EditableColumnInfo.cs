using CommunityToolkit.Mvvm.ComponentModel;

namespace SqlServerTool.Models
{
    /// <summary>構造タブ編集用の列情報（ColumnInfo から生成）</summary>
    public partial class EditableColumnInfo : ObservableObject
    {
        /// <summary>変更前の列名（リネーム検出に使用）</summary>
        public string OriginalName { get; }

        /// <summary>新規追加列かどうか</summary>
        public bool IsNew { get; }

        [ObservableProperty] private string columnName   = string.Empty;
        [ObservableProperty] private string dataType     = string.Empty;
        [ObservableProperty] private bool   isNullable   = true;
        [ObservableProperty] private bool   isPrimaryKey = false;
        [ObservableProperty] private string defaultValue = string.Empty;

        /// <summary>既存列からの生成</summary>
        public EditableColumnInfo(ColumnInfo source)
        {
            OriginalName = source.ColumnName;
            ColumnName   = source.ColumnName;
            DataType     = source.DataType;
            IsNullable   = source.IsNullable;
            IsPrimaryKey = source.IsPrimaryKey;
            DefaultValue = source.DefaultValue;
            IsNew        = false;
        }

        /// <summary>新規列の生成</summary>
        public EditableColumnInfo()
        {
            OriginalName = string.Empty;
            IsNew        = true;
        }
    }
}
