namespace SqlServerTool.Models
{
    public class ScriptOptions
    {
        public string OutputFolder { get; set; } = @"C:\";
        public bool OverwriteExisting { get; set; } = true;
        public bool IncludeDrop { get; set; } = true;
        public bool IncludeData { get; set; } = true;
        public bool IncludeFilegroup { get; set; } = true;
        public bool IncludeViewColumns { get; set; } = true;
    }
}
