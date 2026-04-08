namespace SqlServerTool.Models
{
    public class ColumnInfo
    {
        public string ColumnName { get; set; } = string.Empty;
        public string LogicalName { get; set; } = string.Empty;
        public string DataType { get; set; } = string.Empty;
        public bool IsNullable { get; set; }
        public bool IsPrimaryKey { get; set; }
        public bool IsForeignKey { get; set; }
        public bool IsIdentity { get; set; }
        public string DefaultValue { get; set; } = string.Empty;

        // 表示用
        public string NullableText => IsNullable ? "○" : "×";
        public string PrimaryKeyText => IsPrimaryKey ? "PK" : string.Empty;
        public string ForeignKeyText => IsForeignKey ? "FK" : string.Empty;
    }
}