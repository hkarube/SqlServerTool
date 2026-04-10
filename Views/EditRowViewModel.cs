using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SqlServerTool.Models;
using System.Collections.ObjectModel;

namespace SqlServerTool.ViewModels
{
    public partial class EditRowViewModel : ObservableObject
    {
        public string TableName    { get; }
        public bool   IsNewRecord  { get; init; } = false;
        public ObservableCollection<EditFieldItem> Fields { get; } = new();
        public bool IsConfirmed { get; private set; } = false;

        public EditRowViewModel(string tableName,
            IEnumerable<ColumnInfo> columns,
            IDictionary<string, string> rowData)
        {
            TableName = tableName;
            foreach (var col in columns)
            {
                var value = rowData.ContainsKey(col.ColumnName)
                    ? rowData[col.ColumnName] : string.Empty;
                Fields.Add(new EditFieldItem
                {
                    ColumnName  = col.ColumnName,
                    LogicalName = col.LogicalName,
                    DataType    = col.DataType,
                    IsPrimaryKey = col.IsPrimaryKey,
                    IsIdentity  = col.IsIdentity,
                    Value         = value == "(NULL)" ? string.Empty : value,
                    IsNull        = value == "(NULL)",
                    OriginalValue = value,
                });
            }
        }

        [RelayCommand]
        private void Confirm()
        {
            IsConfirmed = true;
        }
    }

    public partial class EditFieldItem : ObservableObject
    {
        public string ColumnName { get; set; } = string.Empty;
        public string LogicalName { get; set; } = string.Empty;
        public string DataType { get; set; } = string.Empty;
        public bool IsPrimaryKey  { get; set; }
        public bool IsIdentity    { get; set; }
        public bool IsNotIdentity  => !IsIdentity;
        public bool HasLogicalName => !string.IsNullOrEmpty(LogicalName);

        /// <summary>
        /// 画面サイズに合わせて縦方向に伸長すべき型か。
        /// (n)varchar(MAX) / (n)text / xml / image / varbinary(MAX) 等が対象。
        /// </summary>
        public bool IsMaxType =>
            DataType.Contains("MAX", StringComparison.OrdinalIgnoreCase) ||
            DataType.Equals("text",  StringComparison.OrdinalIgnoreCase) ||
            DataType.Equals("ntext", StringComparison.OrdinalIgnoreCase) ||
            DataType.Equals("xml",   StringComparison.OrdinalIgnoreCase) ||
            DataType.Equals("image", StringComparison.OrdinalIgnoreCase);
        public string OriginalValue { get; set; } = string.Empty;

        [ObservableProperty]
        private string value = string.Empty;

        [ObservableProperty]
        private bool isNull = false;

        // NULL チェックボックスONで値入力を無効化
        partial void OnIsNullChanged(bool value)
        {
            if (value) Value = string.Empty;
        }
    }
}