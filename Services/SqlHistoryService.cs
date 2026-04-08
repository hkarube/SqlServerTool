using SqlServerTool.Models;
using System.Collections.ObjectModel;
using System.IO;
using System.Xml.Linq;

namespace SqlServerTool.Services
{
    public class SqlHistoryService
    {
        private static SqlHistoryService? _instance;
        public static SqlHistoryService Instance => _instance ??= new SqlHistoryService();

        private static readonly string HistoryFilePath =
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "SqlServerTool", "SqlHistory.xml");

        public ObservableCollection<SqlHistoryEntry> Entries { get; } = new();

        private SqlHistoryService() => Load();

        public void Add(SqlHistoryEntry entry)
        {
            Entries.Insert(0, entry);
            if (Entries.Count > 1000) Entries.RemoveAt(Entries.Count - 1);
            Save();
        }

        public void Clear()
        {
            Entries.Clear();
            Save();
        }

        private void Save()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(HistoryFilePath)!);
                var doc = new XDocument(
                    new XElement("SqlHistory",
                        Entries.Select(e =>
                            new XElement("Entry",
                                new XAttribute("ExecutedAt",    e.ExecutedAt.ToString("o")),
                                new XAttribute("OperationType", e.OperationType),
                                new XAttribute("ObjectName",    e.ObjectName),
                                new XElement("Sql", e.Sql)))));
                doc.Save(HistoryFilePath);
            }
            catch { /* 保存エラーは無視 */ }
        }

        private void Load()
        {
            if (!File.Exists(HistoryFilePath)) return;
            try
            {
                var doc = XDocument.Load(HistoryFilePath);
                foreach (var el in doc.Root!.Elements("Entry"))
                    Entries.Add(new SqlHistoryEntry
                    {
                        ExecutedAt    = DateTime.Parse(el.Attribute("ExecutedAt")!.Value),
                        OperationType = el.Attribute("OperationType")!.Value,
                        ObjectName    = el.Attribute("ObjectName")!.Value,
                        Sql           = el.Element("Sql")!.Value
                    });
            }
            catch { /* 壊れたファイルは無視 */ }
        }
    }
}
